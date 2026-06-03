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
        public static async Task ExecuterAsync()
        {
            var cfg = BesanconConfig.Charger();

            var rub = EtatApplication.RubidiumActif;
            if (rub == null)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_PAS_RUBIDIUM",
                    "Aucun rubidium actif — tâche Besançon ignorée.");
                return;
            }

            string? contenu = await BesanconFtpService.TelechargerAsync(cfg);
            if (contenu == null) return;

            SauvegarderBrut(contenu);   // copie datée brute (équiv. SavBesancon du legacy)

            var mesures = BesanconParser.Parser(contenu);
            var donnees = BesanconStore.Charger();

            int nouvelles = 0;
            foreach (var m in mesures)
                if (BesanconStore.UpsertValeurJournaliere(donnees, rub.Id, m.Mjd, m.Valeur)) nouvelles++;

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

            BesanconStore.Sauvegarder(donnees);

            Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_OK",
                $"Besançon intégré : {mesures.Count} valeur(s) lue(s), {nouvelles} nouvelle(s) "
              + $"pour le rubidium « {rub.Designation} » (#{rub.Id}).");

            if (cfg.SupprimerApresTelechargement)
                await BesanconFtpService.SupprimerDistantAsync(cfg);
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

        private static void SauvegarderBrut(string contenu)
        {
            try
            {
                string dossier = Path.Combine(CheminsMetrologo.Rubidiums, "SavBesancon");
                Directory.CreateDirectory(dossier);
                File.WriteAllText(Path.Combine(dossier, $"{DateTime.Now:yyyyMMdd_HHmmss}.txt"), contenu);
            }
            catch { /* best-effort */ }
        }
    }
}
