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
    /// Planificateur de la tâche quotidienne Besançon (équivalent du timer <c>tmrBesancon</c> du
    /// legacy) : se déclenche chaque jour à l'heure configurée (défaut 09h50), télécharge le
    /// fichier corrigé par FTP, l'intègre, et le mardi recalcule les moyennes hebdomadaires.
    ///
    /// Ne tourne que si <see cref="BesanconConfig.Active"/> est vrai (à activer sur UN poste).
    /// </summary>
    public static class BesanconScheduler
    {
        private static Timer? _timer;
        private static readonly object _sync = new();

        /// <summary>
        /// Désignation de l'UNIQUE rubidium dont la fréquence de référence est pilotée
        /// automatiquement par la moyenne hebdomadaire Besançon. Pour tout autre rubidium,
        /// la fréquence reste celle saisie manuellement dans le catalogue — on n'y touche pas.
        /// </summary>
        private const string RubidiumPiloteBesancon = "E10-Y8";

        /// <summary>Nombre de semaines écoulées contrôlées lors du rattrapage des moyennes manquantes.</summary>
        private const int NbSemainesRattrapage = 6;

        /// <summary>
        /// Levé après chaque mise à jour du suivi Besançon (récupération quotidienne, rattrapage
        /// d'une moyenne manquante) — l'écran principal s'y abonne pour rafraîchir son voyant +
        /// le rapport texte. Marshalé sur le Dispatcher quand une UI est présente.
        /// </summary>
        public static event EventHandler? StatutChange;

        /// <summary>
        /// Date du jour de référence pour tout le calcul (rattrapage, fenêtres hebdo). Lue sur
        /// l'horloge locale du poste — source unique pour rester cohérente partout dans la classe.
        /// </summary>
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

        /// <summary>
        /// Relit la configuration et l'applique immédiatement : (re)programme la tâche si
        /// <see cref="BesanconConfig.Active"/>, sinon l'arrête. À appeler après modification des
        /// paramètres depuis l'écran Admin pour que le changement prenne effet sans redémarrage.
        /// </summary>
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

        /// <summary>
        /// Exécute la tâche une fois : télécharge le fichier FTP, le parse, et ajoute les
        /// valeurs journalières au fichier texte cumulatif (<see cref="BesanconTxtStore"/>) —
        /// AUCUNE écriture en base SQL.
        ///
        /// <para/>Si <paramref name="forcer"/> est faux (déclenchement automatique quotidien) et
        /// qu'une récupération a DÉJÀ abouti aujourd'hui sur n'importe quel poste (marqueur
        /// partagé), le téléchargement est ignoré — évite que plusieurs postes retéléchargent le
        /// même fichier. Le déclenchement manuel (« Forcer ») passe <c>true</c> et ignore le garde-fou.
        /// </summary>
        public static async Task<ResultatBesancon> ExecuterAsync(bool forcer = false)
        {
            var res = new ResultatBesancon { Destination = BesanconTxtStore.CheminValeurs };

            // Garde-fou multi-poste : déjà fait aujourd'hui ? → on n'y retouche pas (sauf « Forcer »).
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

            // Ajout au fichier texte cumulatif (sans doublon de MJD). Toute erreur d'écriture
            // est remontée à l'admin — mais on ne touche plus à aucune base SQL.
            int maxMjd = 0;
            try
            {
                res.Nouvelles = await BesanconTxtStore.AjouterAsync(mesures);
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

            // Notifie l'écran d'accueil (singleton, abonné à StatutChange) pour qu'il relise le
            // fichier txt et rafraîchisse le voyant + le rapport sans attendre un redémarrage.
            NotifierStatutChange();

            Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_OK",
                $"Besançon récupéré : {mesures.Count} valeur(s) lue(s), {res.Nouvelles} nouvelle(s) ajoutée(s) "
              + $"au fichier {BesanconTxtStore.CheminValeurs}. Brut : {res.CheminBrut ?? "NON ÉCRIT"}.");

            if (cfg.SupprimerApresTelechargement)
                await BesanconFtpService.SupprimerDistantAsync(cfg);

            return res;
        }

        /// <summary>
        /// Lève <see cref="StatutChange"/> (marshalé sur le Dispatcher si une UI est présente)
        /// pour que l'écran d'accueil — un singleton créé au démarrage — relise le fichier txt et
        /// rafraîchisse le voyant + le rapport après une récupération (manuelle ou quotidienne).
        /// </summary>
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

        private static async Task TenterCalculHebdoAsync(int rubId, bool avecGps, int mardiMjd)
        {
            var r = await BesanconStore.CalculerMoyenneHebdoAsync(rubId, mardiMjd);
            if (r.HasValue)
            {
                await BesanconStore.UpsertMoyenneHebdoAsync(rubId, avecGps, mardiMjd, r.Value.moyenne, r.Value.deltaTps);
                Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_HEBDO",
                    $"Moyenne hebdo rubidium #{rubId} (mardi MJD {mardiMjd}) : {r.Value.moyenne:G9} "
                  + $"(delta {r.Value.deltaTps:G9} s/jour).");
            }
        }

        /// <summary>
        /// Vérifie les dernières semaines écoulées et calcule à la volée toute moyenne hebdo
        /// manquante DONT les 7 valeurs journalières sont déjà disponibles — typiquement au
        /// démarrage de l'app (« est-ce que le calcul de la semaine passée a été fait ? sinon
        /// fais-le tout de suite »). Si les 7 jours ne sont pas encore là (ex. on est lundi),
        /// rien n'est calculé : la tâche quotidienne le fera dès que les valeurs arriveront.
        ///
        /// <para/>Idempotent (n'écrase pas une moyenne existante) et sans coût FTP — pur calcul
        /// SQL. Réinjecte la dernière moyenne dans le rubidium piloté (E10-Y8) si une semaine
        /// vient d'être rattrapée, puis notifie l'écran principal.
        /// </summary>
        public static async Task AssurerCalculsHebdoManquantsAsync()
        {
            try
            {
                var rub = EtatApplication.RubidiumActif;
                if (rub == null) return;

                var today = Aujourdhui;
                int todayMjd = JourJulien.VersMjd(today);
                DateTime mardi = DernierMardiInclus(today);
                bool rattrape = false;

                for (int k = 0; k < NbSemainesRattrapage; k++)
                {
                    int mardiMjd = JourJulien.VersMjd(mardi);
                    int fin = mardiMjd - 1;
                    // Semaine entièrement écoulée et moyenne absente → on tente le calcul.
                    if (fin < todayMjd && !await BesanconStore.MoyenneHebdoExisteAsync(rub.Id, mardiMjd))
                    {
                        await TenterCalculHebdoAsync(rub.Id, rub.AvecGPS, mardiMjd);
                        if (await BesanconStore.MoyenneHebdoExisteAsync(rub.Id, mardiMjd))
                            rattrape = true;
                    }
                    mardi = mardi.AddDays(-7);
                }

                // Réinjecte la dernière moyenne disponible (utile si une semaine a été rattrapée).
                var dh = await BesanconStore.DerniereMoyenneHebdoAsync(rub.Id);
                if (dh.HasValue)
                    InjecterReferenceDansRubidiumActif(rub, dh.Value.moyenne);

                if (rattrape)
                {
                    Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_RATTRAPAGE",
                        "Moyenne(s) hebdomadaire(s) manquante(s) recalculée(s) au démarrage.");
                    await GenererRapportEtNotifierAsync(rub);
                }
            }
            catch (Exception ex)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_RATTRAPAGE_KO",
                    $"Rattrapage des moyennes hebdo échoué : {ex.Message}");
            }
        }

        /// <summary>Régénère le rapport texte de suivi puis lève <see cref="StatutChange"/> (sur le Dispatcher si UI présente).</summary>
        private static async Task GenererRapportEtNotifierAsync(Rubidium rub)
        {
            try { await BesanconSuiviService.EvaluerAsync(rub, Aujourdhui); }
            catch { /* best-effort : le panneau réévaluera de son côté */ }

            var disp = System.Windows.Application.Current?.Dispatcher;
            if (disp != null)
                disp.BeginInvoke(new Action(() => StatutChange?.Invoke(null, EventArgs.Empty)));
            else
                StatutChange?.Invoke(null, EventArgs.Empty);
        }

        /// <summary>Le mardi de la semaine courante (ou aujourd'hui si l'on est un mardi).</summary>
        private static DateTime DernierMardiInclus(DateTime d)
        {
            int diff = ((int)d.DayOfWeek - (int)DayOfWeek.Tuesday + 7) % 7;
            return d.Date.AddDays(-diff);
        }

        /// <summary>
        /// Reporte la moyenne hebdomadaire Besançon comme fréquence de référence du rubidium
        /// actif — mais SEULEMENT s'il s'agit du rubidium <see cref="RubidiumPiloteBesancon"/>
        /// (« E10-Y8 »). La moyenne (valeurs déjà corrigées du fichier <c>ef_utcop</c>, rapportées
        /// au 10 MHz) devient directement la <see cref="Rubidium.FrequenceMoyenne"/> utilisée par
        /// les mesures (zone <c>ZNFreqRef</c> → <c>Cal_freq_corrigee</c>).
        ///
        /// <para/>Persiste sur le partage (<c>rubidium-actif.json</c> + catalogue) puis notifie
        /// l'UI sur le thread Dispatcher (la tâche tourne en arrière-plan via <see cref="Timer"/>,
        /// on ne lève donc pas l'événement directement sur ce thread).
        /// </summary>
        private static void InjecterReferenceDansRubidiumActif(Rubidium rub, double moyenne)
        {
            // Cadré sur un seul rubidium : aucune autre référence n'est modifiée.
            if (!string.Equals(rub.Designation?.Trim(), RubidiumPiloteBesancon,
                    StringComparison.OrdinalIgnoreCase))
                return;

            // Garde-fou : on n'écrase jamais la référence par une valeur non plausible
            // (fichier incomplet, parsing partiel…). La fréquence de référence est ~10 MHz.
            if (!(moyenne > 0))
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_REF_IGNOREE",
                    $"Moyenne hebdo non plausible ({moyenne:G9} Hz) — fréquence de référence de "
                  + $"« {rub.Designation} » laissée inchangée.");
                return;
            }

            double ancienne = rub.FrequenceMoyenne;
            if (Math.Abs(ancienne - moyenne) < 1e-9) return;   // déjà à jour : rien à faire

            // 1. Mise à jour en mémoire (rub == EtatApplication.RubidiumActif, même référence)
            //    → la prochaine mesure lira directement la nouvelle valeur.
            rub.FrequenceMoyenne = moyenne;

            // 2. Persistance partagée (rubidium-actif.json) + repli local.
            Models.Preferences.SauvegarderRubidium(rub);

            // 3. Répercute aussi dans le catalogue partagé pour que la valeur survive à une
            //    re-sélection ultérieure du rubidium depuis la liste.
            try
            {
                var catalogue = Models.Preferences.CatalogueRubidiums.ToList();
                var cible = catalogue.FirstOrDefault(r =>
                    string.Equals(r.Designation?.Trim(), RubidiumPiloteBesancon,
                        StringComparison.OrdinalIgnoreCase));
                if (cible != null && Math.Abs(cible.FrequenceMoyenne - moyenne) > 1e-9)
                {
                    cible.FrequenceMoyenne = moyenne;
                    Models.Preferences.SauvegarderCatalogueRubidiums(catalogue);
                }
            }
            catch (Exception ex)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_REF_CATALOGUE_KO",
                    $"Mise à jour du catalogue pour « {rub.Designation} » échouée : {ex.Message}");
            }

            // 4. Notifie l'UI (barre de statut, écrans) sur le thread Dispatcher. En contexte
            //    headless/test (pas d'Application WPF), la persistance ci-dessus suffit.
            var disp = System.Windows.Application.Current?.Dispatcher;
            if (disp != null)
                disp.BeginInvoke(new Action(() => EtatApplication.NotifierRubidiumActifChange()));

            Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_REF_INJECTEE",
                $"Fréquence de référence de « {rub.Designation} » mise à jour depuis la moyenne "
              + $"hebdo Besançon : {ancienne:G9} Hz → {moyenne:G9} Hz.");
        }

        /// <summary>
        /// Dépose une copie datée du fichier brut sur le partage (dossier <c>SavBesancon</c>),
        /// pour qu'il reste consultable tel quel. Retourne le chemin écrit, ou <c>null</c> si
        /// l'écriture a échoué — auquel cas le chemin tenté et l'erreur sont loggués (utile pour
        /// diagnostiquer un partage réseau injoignable).
        /// </summary>
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

    /// <summary>
    /// Compte-rendu d'une exécution de la tâche Besançon — sert à afficher à l'admin un résumé
    /// concret (fichier récupéré, chemins exacts sur le partage, valeurs lues/intégrées, dernière
    /// moyenne hebdo) ou la cause précise d'échec, au lieu d'un simple « voir le Journal ».
    /// </summary>
    public sealed class ResultatBesancon
    {
        public bool Telecharge { get; set; }
        /// <summary>Vrai si la récupération a été ignorée car déjà faite aujourd'hui (marqueur partagé).</summary>
        public bool DejaFait { get; set; }
        public string? CheminBrut { get; set; }
        /// <summary>Où le suivi est enregistré (ex. « Base SQL BASE_E2M »).</summary>
        public string Destination { get; set; } = "";
        /// <summary>Vrai si l'enregistrement en base a réussi.</summary>
        public bool EnregistrementOk { get; set; }
        public int ValeursLues { get; set; }
        public int Nouvelles { get; set; }
        public int TotalJournalieres { get; set; }
        public double? DerniereMoyenneHebdo { get; set; }
        public int DerniereMoyenneHebdoMjd { get; set; }
        public string RubidiumDesignation { get; set; } = "";
        public string? Erreur { get; set; }

        public bool Succes => Telecharge && Erreur == null;
    }
}
