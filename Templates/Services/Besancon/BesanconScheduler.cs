using System;
using System.IO;
using System.Linq;
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
        /// Exécute la tâche une fois : télécharge le fichier FTP, le parse, intègre les valeurs
        /// journalières pour le rubidium actif, et (le mardi) calcule les moyennes hebdomadaires.
        /// Public pour permettre un déclenchement manuel (« Forcer la récupération »).
        /// </summary>
        public static async Task<ResultatBesancon> ExecuterAsync()
        {
            var res = new ResultatBesancon { CheminJson = BesanconStore.Chemin };
            var cfg = BesanconConfig.Charger();

            var rub = EtatApplication.RubidiumActif;
            if (rub == null)
            {
                res.Erreur = "Aucun rubidium actif — sélectionne un rubidium avant de récupérer Besançon.";
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_PAS_RUBIDIUM", res.Erreur);
                return res;
            }
            res.RubidiumDesignation = rub.Designation;

            string? contenu = await BesanconFtpService.TelechargerAsync(cfg);
            if (contenu == null)
            {
                res.Erreur = $"Téléchargement FTP échoué ou FTP non configuré "
                           + $"(hôte « {cfg.FtpHote} », fichier « {cfg.FichierDistant} »). Voir le Journal (Système).";
                return res;
            }
            res.Telecharge = true;

            // Dépose le brut sur le partage (récupère le chemin exact, ou null si échec).
            res.CheminBrut = SauvegarderBrut(contenu);

            var mesures = BesanconParser.Parser(contenu);
            res.ValeursLues = mesures.Count;

            var donnees = BesanconStore.Charger();
            int nouvelles = 0;
            foreach (var m in mesures)
                if (BesanconStore.UpsertValeurJournaliere(donnees, rub.Id, m.Mjd, m.Valeur)) nouvelles++;
            res.Nouvelles = nouvelles;

            // Moyennes hebdo : le mardi, on (re)calcule les 3 mardis précédents + le mardi courant
            // (comme le legacy) ; les autres jours, on tente le dernier mardi passé.
            var aujourdhui = DateTime.Today;
            if (aujourdhui.DayOfWeek == DayOfWeek.Tuesday)
            {
                int mjd = JourJulien.VersMjd(aujourdhui);
                foreach (int mardi in new[] { mjd - 21, mjd - 14, mjd - 7, mjd })
                    TenterCalculHebdo(donnees, rub.Id, mardi);
            }
            else
            {
                int decalDepuisMardi = ((int)aujourdhui.DayOfWeek - (int)DayOfWeek.Tuesday + 7) % 7;
                int mjdDernierMardi = JourJulien.VersMjd(aujourdhui.AddDays(-decalDepuisMardi));
                TenterCalculHebdo(donnees, rub.Id, mjdDernierMardi);
            }

            res.SauvegardeJsonOk = BesanconStore.Sauvegarder(donnees);

            // Synthèse pour la consultation immédiate (stock total + dernière moyenne hebdo).
            res.TotalJournalieres = donnees.Journalieres.Count(v => v.RubidiumId == rub.Id);
            var derniereHebdo = donnees.Hebdos
                .Where(h => h.RubidiumId == rub.Id)
                .OrderByDescending(h => h.MardiMjd)
                .FirstOrDefault();
            if (derniereHebdo != null)
            {
                res.DerniereMoyenneHebdo = derniereHebdo.Moyenne;
                res.DerniereMoyenneHebdoMjd = derniereHebdo.MardiMjd;
            }

            Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_OK",
                $"Besançon intégré : {mesures.Count} valeur(s) lue(s), {nouvelles} nouvelle(s) "
              + $"pour le rubidium « {rub.Designation} » (#{rub.Id}). "
              + $"Brut : {res.CheminBrut ?? "NON ÉCRIT"} · JSON : {(res.SauvegardeJsonOk ? res.CheminJson : "NON ÉCRIT")}.");

            if (cfg.SupprimerApresTelechargement)
                await BesanconFtpService.SupprimerDistantAsync(cfg);

            return res;
        }

        private static void TenterCalculHebdo(DonneesBesancon d, int rubId, int mardiMjd)
        {
            if (BesanconStore.CalculerMoyenneHebdo(d, rubId, mardiMjd, out double moyenne, out double delta))
            {
                BesanconStore.UpsertMoyenneHebdo(d, rubId, mardiMjd, moyenne, delta);
                Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_HEBDO",
                    $"Moyenne hebdo rubidium #{rubId} (mardi MJD {mardiMjd}) : {moyenne:G9} "
                  + $"(delta {delta:G9} s/jour).");
            }
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
        public string? CheminBrut { get; set; }
        public string CheminJson { get; set; } = "";
        public bool SauvegardeJsonOk { get; set; }
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
