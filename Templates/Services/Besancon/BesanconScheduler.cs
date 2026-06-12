using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Metrologo.Models;
using Metrologo.Services.Journal;

namespace Metrologo.Services.Besancon
{
    /// <summary>
    /// Tâche quotidienne Besançon (reprend le timer tmrBesancon du legacy) : chaque jour à l'heure
    /// configurée, télécharge le fichier corrigé par FTP, l'intègre, et le mardi recalcule la moyenne
    /// hebdo. Ne tourne que si la config est active sur le poste.
    /// </summary>
    public static class BesanconScheduler
    {
        private static Timer? _timer;
        private static readonly object _sync = new();

        /// <summary>Base 10 MHz sur laquelle l'écart hebdo s'applique : freq = 10 MHz * (1 + écart).</summary>
        private const double FrequenceNominaleHz = 10_000_000.0;

        /// <summary>En dessous de ce seuil (Hz) la fréquence est considérée inchangée, pour éviter
        /// les réécritures inutiles (bien plus fin que les offsets hebdo ~1e-7 Hz).</summary>
        private const double SeuilHz = 1e-9;

        /// <summary>Si la nouvelle valeur hebdo passe sous cette limite, on alerte l'opérateur par pop-up.</summary>
        private const double SeuilAlerteHebdo = 1e-13;

        /// <summary>Levé après chaque mise à jour du suivi Besançon ; l'écran principal s'y abonne
        /// pour rafraîchir voyant et rapport. Marshalé sur le Dispatcher si une UI est présente.</summary>
        public static event EventHandler? StatutChange;

        /// <summary>
        /// Demande une pop-up sur CE poste (titre, message, avertissement). La tâche ne s'exécute que
        /// sur un poste par jour et le watcher admin ne notifie que les AUTRES postes : on combine donc
        /// pop-up directe ici + entrée au journal d'audit pour le reste du parc.
        /// </summary>
        public static event Action<string, string, bool>? PopupRubidiumDemandee;

        /// <summary>Lève PopupRubidiumDemandee sur le thread UI si une UI est présente.</summary>
        private static void NotifierPopupRubidium(string titre, string message, bool avertissement)
        {
            var disp = System.Windows.Application.Current?.Dispatcher;
            if (disp != null)
                disp.BeginInvoke(new Action(() => PopupRubidiumDemandee?.Invoke(titre, message, avertissement)));
            else
                PopupRubidiumDemandee?.Invoke(titre, message, avertissement);
        }

        /// <summary>Jour de référence de tous les calculs (horloge locale), centralisé ici pour rester cohérent.</summary>
        private static DateTime Aujourdhui => DateTime.Today;

        /// <summary>Démarre la planification (no-op si la tâche est désactivée sur ce poste).</summary>
        public static void Demarrer()
        {
            var cfg = BesanconConfig.Charger();
            if (!cfg.Active)
            {
                Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_INACTIF",
                    "Tâche Besançon désactivée sur ce poste (besancon.ftp.json : Active=false).");
                return;
            }
            Programmer(cfg);

            // Rattrapage au démarrage : le timer ne sert à rien si l'app n'était pas ouverte à
            // l'heure prévue. Donc si on lance Metrologo après l'heure de déclenchement et que la
            // récupération du jour n'a pas été faite (marqueur partagé), on la lance maintenant.
            // On ne rattrape qu'une fois l'heure passée (le fichier Besançon n'est publié qu'à ce
            // moment-là), et ExecuterAsync(forcer:false) respecte DejaRecupereAujourdhui() donc
            // pas de doublon si un autre poste l'a déjà fait. Fire-and-forget.
            if (DateTime.Now.TimeOfDay >= cfg.HeureParsee())
                _ = RattrapageDemarrageAsync();
        }

        private static async Task RattrapageDemarrageAsync()
        {
            try
            {
                var res = await ExecuterAsync();
                if (!res.DejaFait)
                    Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_RATTRAPAGE",
                        "Récupération Besançon rattrapée au démarrage (heure de déclenchement déjà passée).");
            }
            catch (Exception ex)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_RATTRAPAGE_KO",
                    $"Rattrapage Besançon au démarrage échoué : {ex.Message}");
            }
        }

        /// <summary>Arrête la planification quotidienne (dispose le timer s'il tourne).</summary>
        public static void Arreter()
        {
            lock (_sync)
            {
                _timer?.Dispose();
                _timer = null;
            }
        }

        /// <summary>Relit la config et l'applique tout de suite (à appeler après modification
        /// depuis l'écran Admin, évite un redémarrage).</summary>
        public static void Reconfigurer()
        {
            Arreter();
            Demarrer();
        }

        private static void Programmer(BesanconConfig cfg)
        {
            var heure = cfg.HeureParsee();
            var maintenant = DateTime.Now;
            var prochain = DateTime.Today.Add(heure);
            if (prochain <= maintenant) prochain = prochain.AddDays(1);
            var delai = prochain - maintenant;

            lock (_sync)
            {
                _timer?.Dispose();
                _timer = new Timer(_ => _ = ExecuterEtReprogrammerAsync(), null, delai, Timeout.InfiniteTimeSpan);
            }
            Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_PROGRAMME",
                $"Prochaine récupération Besançon programmée le {prochain:dd/MM/yyyy à HH:mm}.");
        }

        private static async Task ExecuterEtReprogrammerAsync()
        {
            try { await ExecuterAsync(); }
            catch (Exception ex)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_TACHE_KO",
                    $"Tâche Besançon échouée : {ex.Message}");
            }
            finally
            {
                // Reprogramme pour le lendemain à la même heure (re-lit la config au cas où).
                Programmer(BesanconConfig.Charger());
            }
        }

        /// <summary>Exécute la tâche une fois : téléchargement FTP, parsing, ajout au fichier txt
        /// cumulatif (aucune écriture SQL). Sans forcer, no-op si un poste a déjà récupéré
        /// aujourd'hui (marqueur partagé) ; forcer=true (bouton Forcer) ignore ce garde-fou.</summary>
        public static async Task<ResultatBesancon> ExecuterAsync(bool forcer = false)
        {
            var res = new ResultatBesancon { Destination = BesanconTxtStore.CheminValeurs };

            // garde-fou multi-poste : déjà fait aujourd'hui, on ne refait pas (sauf forçage manuel)
            if (!forcer && DejaRecupereAujourdhui())
            {
                res.DejaFait = true;
                Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_DEJA_FAIT",
                    "Récupération Besançon déjà effectuée aujourd'hui (marqueur partagé) — ignorée sur ce poste.");
                return res;
            }

            var cfg = BesanconConfig.Charger();

            var ftp = await BesanconFtpService.TelechargerAsync(cfg);
            if (!ftp.Ok)
            {
                res.Erreur = ftp.ConfigManquante
                    ? $"FTP non configuré : {ftp.Erreur}"
                    : $"Échec FTP sur {ftp.Url}\n→ {ftp.Erreur}";
                return res;
            }
            res.Telecharge = true;
            string contenu = ftp.Contenu!;

            // Dépose le brut sur le partage (récupère le chemin exact, ou null si échec).
            res.CheminBrut = SauvegarderBrut(contenu);

            var mesures = BesanconParser.Parser(contenu);
            res.ValeursLues = mesures.Count;

            // ajout au fichier texte cumulatif (sans doublon de MJD) ; toute erreur d'écriture
            // est remontée à l'admin, mais on ne touche plus à aucune base SQL
            int maxMjd = 0;
            try
            {
                var ajout = await BesanconTxtStore.AjouterAsync(mesures);
                res.Nouvelles = ajout.Nouvelles;
                res.Corrections = ajout.Corrections.Count;

                // Corrections rares mais importantes (Besançon révise parfois une valeur passée) :
                // elles modifient l'historique → potentiellement la moyenne hebdo → la fréquence de
                // référence injectée plus bas. On les trace explicitement (date, ancienne → nouvelle).
                if (ajout.Corrections.Count > 0)
                {
                    string details = string.Join(" ; ", ajout.Corrections.Select(c =>
                        $"{JourJulien.DepuisMjd(c.Mjd):dd/MM/yyyy} {c.Ancienne:G9}→{c.Nouvelle:G9}"));
                    Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_CORRECTION",
                        $"{ajout.Corrections.Count} valeur(s) passée(s) corrigée(s) par Besançon : {details}");
                }

                var toutes = await BesanconTxtStore.LireAsync();
                res.TotalJournalieres = toutes.Count;
                if (toutes.Count > 0) maxMjd = toutes.Keys.Max();
                res.EnregistrementOk = true;
            }
            catch (Exception exTxt)
            {
                res.Erreur = $"Écriture du fichier txt cumulatif échouée ({BesanconTxtStore.CheminValeurs}) : {exTxt.Message}";
                Journal.Journal.Erreur(CategorieLog.Systeme, "BESANCON_TXT_KO", res.Erreur);
                return res;
            }

            // Marqueur partagé : signale aux autres postes que c'est fait pour aujourd'hui.
            EcrireMarqueur(maxMjd);

            // Injecte la fréquence de référence du rubidium piloté E10-Y8 à partir de l'écart de la
            // dernière semaine COMPLÈTE mardi→lundi : 10 MHz × (1 + écart). Recalculé à chaque
            // récupération quotidienne ; ne change effectivement que le mardi (rotation de semaine).
            try
            {
                var ecartHebdo = await BesanconSuiviService.EcartHebdoCompletAsync(DateTime.Today);
                if (ecartHebdo.HasValue)
                {
                    bool valeurChangee = InjecterReferenceRubidiumActif(ecartHebdo.Value);

                    // alerte quand la nouvelle valeur hebdo passe sous la limite de 1e-13 ;
                    // conditionnée au changement effectif (rotation de semaine) pour ne pas
                    // répéter l'alerte chaque jour tant que la valeur ne bouge pas
                    if (valeurChangee && ecartHebdo.Value < SeuilAlerteHebdo)
                    {
                        string detail = $"L'écart hebdomadaire Besançon ({ecartHebdo.Value:G6}) est passé "
                                      + $"sous la limite de {SeuilAlerteHebdo:G1}.";
                        // catégorie Rubidium + action en liste blanche : passe au journal d'audit,
                        // donc les autres postes sont avertis ; pop-up directe ci-dessous pour CE poste
                        Journal.Journal.Warn(CategorieLog.Rubidium, "RUBIDIUM_HEBDO_SOUS_SEUIL", detail);
                        NotifierPopupRubidium("⚠ Écart hebdo sous la limite",
                            detail + "\n\nVérifie l'étalonnage / la chaîne de référence du rubidium.",
                            avertissement: true);
                    }
                }
            }
            catch (Exception exRef)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_REF_KO",
                    $"Injection de la fréquence de référence E10-Y8 échouée : {exRef.Message}");
            }

            // Notifie l'écran d'accueil (singleton, abonné à StatutChange) pour qu'il relise le
            // fichier txt et rafraîchisse le voyant + le rapport sans attendre un redémarrage.
            NotifierStatutChange();

            Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_OK",
                $"Besançon récupéré : {mesures.Count} valeur(s) lue(s), {res.Nouvelles} nouvelle(s) ajoutée(s), "
              + $"{res.Corrections} correction(s) appliquée(s) au fichier {BesanconTxtStore.CheminValeurs}. "
              + $"Brut : {res.CheminBrut ?? "NON ÉCRIT"}.");

            if (cfg.SupprimerApresTelechargement)
                await BesanconFtpService.SupprimerDistantAsync(cfg);

            return res;
        }

        /// <summary>Lève StatutChange (sur le Dispatcher si UI) pour que l'écran d'accueil relise
        /// le fichier txt et rafraîchisse voyant + rapport après une récupération.</summary>
        private static void NotifierStatutChange()
        {
            var disp = System.Windows.Application.Current?.Dispatcher;
            if (disp != null)
                disp.BeginInvoke(new Action(() => StatutChange?.Invoke(null, EventArgs.Empty)));
            else
                StatutChange?.Invoke(null, EventArgs.Empty);
        }

        /// <summary>Marqueur PARTAGÉ de la dernière récupération aboutie (date + poste + MJD max).</summary>
        private static string CheminMarqueur =>
            Path.Combine(CheminsMetrologo.Besancon, "derniere_recuperation.json");

        private sealed class MarqueurRecuperation
        {
            public string Date { get; set; } = "";        // jour local, format yyyy-MM-dd
            public string Poste { get; set; } = "";
            public int MaxMjd { get; set; }
            public string Horodatage { get; set; } = "";
        }

        /// <summary>
        /// Vrai si une récupération a DÉJÀ abouti aujourd'hui (n'importe quel poste), d'après le
        /// marqueur partagé. Marqueur absent ou illisible → false (on ne bloque pas la récupération).
        /// </summary>
        private static bool DejaRecupereAujourdhui()
        {
            try
            {
                if (!File.Exists(CheminMarqueur)) return false;
                var m = JsonSerializer.Deserialize<MarqueurRecuperation>(File.ReadAllText(CheminMarqueur));
                return m != null && m.Date == Aujourdhui.ToString("yyyy-MM-dd");
            }
            catch { return false; }
        }

        /// <summary>Écrit le marqueur partagé (best-effort) après une récupération aboutie.</summary>
        private static void EcrireMarqueur(int maxMjd)
        {
            try
            {
                Directory.CreateDirectory(CheminsMetrologo.Besancon);
                var m = new MarqueurRecuperation
                {
                    Date = Aujourdhui.ToString("yyyy-MM-dd"),
                    Poste = Environment.MachineName,
                    MaxMjd = maxMjd,
                    Horodatage = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                };
                File.WriteAllText(CheminMarqueur,
                    JsonSerializer.Serialize(m, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch (Exception ex)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_MARQUEUR_KO",
                    $"Écriture du marqueur de récupération Besançon échouée : {ex.Message}");
            }
        }

        /// <summary>Applique l'écart hebdo Besançon (signé) comme fréquence de référence du rubidium
        /// actif : 10 MHz x (1 + écart). Met à jour l'objet en mémoire, rubidium-actif.json et le
        /// catalogue partagé, puis notifie l'UI. Un rubidium en réglage manuel n'est pas écrasé.
        /// Retourne true seulement si la valeur a réellement changé (rotation de semaine).</summary>
        private static bool InjecterReferenceRubidiumActif(double ecart)
        {
            // garde-fou : écart fini et plausible (les écarts Besançon font ~1e-11..1e-14, |v| <= 1e-9).
            // Les valeurs négatives sont acceptées (référence sous 10 MHz), seul l'invraisemblable est rejeté.
            if (double.IsNaN(ecart) || double.IsInfinity(ecart) || Math.Abs(ecart) > 1e-9)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_REF_IGNOREE",
                    $"Écart hebdo non plausible ({ecart:G9}) — fréquence de référence laissée inchangée.");
                return false;
            }

            var actif = EtatApplication.RubidiumActif;
            if (actif == null)
            {
                Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_REF_SANS_RUBIDIUM",
                    "Aucun rubidium actif — injection de la fréquence de référence Besançon différée.");
                return false;
            }

            // automatique, aucun raccord manuel requis ; seul un rubidium en réglage MANUEL
            // (fréquence saisie explicitement par l'utilisateur) n'est pas écrasé
            if (actif.EstReglageManuel)
            {
                Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_REF_MANUEL",
                    "Rubidium actif en réglage manuel — fréquence de référence laissée inchangée.");
                return false;
            }

            double nouvelleFreq = FrequenceNominaleHz * (1.0 + ecart);
            if (Math.Abs(actif.FrequenceMoyenne - nouvelleFreq) <= SeuilHz)
            {
                Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_REF_INCHANGEE",
                    $"Fréquence de référence de « {actif.Designation} » déjà à jour ({nouvelleFreq:F9} Hz, écart {ecart:G9}).");
                return false;
            }

            double ancienne = actif.FrequenceMoyenne;

            // 1. Mise à jour en mémoire (objet == EtatApplication.RubidiumActif) + persistance partagée.
            actif.FrequenceMoyenne = nouvelleFreq;
            Models.Preferences.SauvegarderRubidium(actif);

            // 2. Répercute dans l'entrée correspondante du catalogue partagé (survit à une re-sélection).
            try
            {
                var catalogue = Models.Preferences.CatalogueRubidiums.ToList();
                var cible = catalogue.FirstOrDefault(r =>
                    string.Equals(r.Designation?.Trim(), actif.Designation?.Trim(), StringComparison.OrdinalIgnoreCase));
                if (cible != null && Math.Abs(cible.FrequenceMoyenne - nouvelleFreq) > SeuilHz)
                {
                    cible.FrequenceMoyenne = nouvelleFreq;
                    Models.Preferences.SauvegarderCatalogueRubidiums(catalogue);
                }
            }
            catch (Exception ex)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_REF_CATALOGUE_KO",
                    $"Mise à jour du catalogue pour « {actif.Designation} » échouée : {ex.Message}");
            }

            // 3. Notifie l'UI (bandeau bas, écrans) sur le thread Dispatcher.
            var disp = System.Windows.Application.Current?.Dispatcher;
            if (disp != null)
                disp.BeginInvoke(new Action(() => EtatApplication.NotifierRubidiumActifChange()));

            Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_REF_INJECTEE",
                $"Fréquence de référence de « {actif.Designation} » mise à jour depuis l'écart hebdo "
              + $"Besançon ({ecart:G9}) : {ancienne:F9} Hz → {nouvelleFreq:F9} Hz.");

            // 4. Trace d'audit (catégorie Rubidium + action en liste blanche) : les autres postes
            //    sont avertis et peuvent recharger la config (rubidium-actif.json mis à jour).
            Journal.Journal.Info(CategorieLog.Rubidium, "RUBIDIUM_VALEUR_MAJ",
                $"« {actif.Designation} » : {ancienne:F6} Hz → {nouvelleFreq:F6} Hz (écart hebdo {ecart:G6}).");

            // 5. Pop-up directe sur CE poste (le watcher d'audit ne notifie que les autres).
            NotifierPopupRubidium("Rubidium — nouvelle valeur",
                $"La fréquence de référence du rubidium « {actif.Designation} » a été mise à jour :\n\n"
              + $"{ancienne:F6} Hz  →  {nouvelleFreq:F6} Hz\n"
              + $"(écart hebdomadaire Besançon : {ecart:G6})\n\n"
              + "Les écrans utilisent désormais cette nouvelle valeur.",
                avertissement: false);

            return true;
        }

        /// <summary>Dépose une copie datée du fichier brut dans SavBesancon sur le partage.
        /// Retourne le chemin écrit, ou null si l'écriture a échoué (chemin et erreur loggués).</summary>
        private static string? SauvegarderBrut(string contenu)
        {
            string dossier = Path.Combine(CheminsMetrologo.Besancon, "SavBesancon");
            string chemin = Path.Combine(dossier, $"{DateTime.Now:yyyyMMdd_HHmmss}.txt");
            try
            {
                Directory.CreateDirectory(dossier);
                File.WriteAllText(chemin, contenu);
                Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_BRUT_OK",
                    $"Fichier brut Besançon déposé sur le partage : {chemin}");
                return chemin;
            }
            catch (Exception ex)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_BRUT_KO",
                    $"Écriture du fichier brut Besançon échouée ({chemin}) : {ex.Message}");
                return null;
            }
        }
    }

    /// <summary>Compte-rendu d'une exécution de la tâche Besançon : résumé concret pour l'admin
    /// (chemins, valeurs lues/intégrées, moyenne hebdo) ou cause précise d'échec.</summary>
    public sealed class ResultatBesancon
    {
        public bool Telecharge { get; set; }
        /// <summary>Vrai si la récupération a été ignorée car déjà faite aujourd'hui (marqueur partagé).</summary>
        public bool DejaFait { get; set; }
        public string? CheminBrut { get; set; }
        /// <summary>Où le suivi est enregistré (chemin du fichier txt cumulatif).</summary>
        public string Destination { get; set; } = "";
        /// <summary>Vrai si l'écriture du fichier txt cumulatif a réussi.</summary>
        public bool EnregistrementOk { get; set; }
        public int ValeursLues { get; set; }
        public int Nouvelles { get; set; }
        /// <summary>Nombre de valeurs passées corrigées par la source (révisions Besançon).</summary>
        public int Corrections { get; set; }
        public int TotalJournalieres { get; set; }
        public double? DerniereMoyenneHebdo { get; set; }
        public int DerniereMoyenneHebdoMjd { get; set; }
        public string RubidiumDesignation { get; set; } = "";
        public string? Erreur { get; set; }

        public bool Succes => Telecharge && Erreur == null;
    }
}
