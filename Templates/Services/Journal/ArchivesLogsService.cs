using System.IO;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Archive les logs du mois écoulé sur le partage réseau : au premier lancement du mois (tout poste),
    /// <c>entries-&lt;moisPrec&gt;.jsonl</c> et <c>.txt</c> passent de <c>M:\Logs\</c> à
    /// <c>M:\Logs\Archives\&lt;moisPrec&gt;\</c>. Les logs du mois en cours restent dans <c>M:\Logs\</c>.
    ///
    /// Multi-postes : verrou fichier exclusif (<c>archive-YYYY-MM.lock</c>) et marqueur de fin
    /// (<c>Archives\YYYY-MM\.done</c>) : un seul poste archive, les autres passent leur chemin.
    /// Si M:\ est inaccessible, rien n'est fait.
    /// </summary>
    public static class ArchivesLogsService
    {
        /// <summary>Le dossier Archives, situé à l'intérieur de Logs (sur le partage).</summary>
        public static string DossierArchivesRacine => Path.Combine(CheminsMetrologo.Logs, "Archives");

        /// <summary>Archive à la demande le mois indiqué (utilisé par le visualiseur de journal).</summary>
        public static Task<int> ArchiverMoisAsync(DateTime mois)
        {
            int n = ArchiverMoisVersDossierInterne(mois);
            return Task.FromResult(n);
        }

        /// <summary>
        /// Au démarrage de l'app : archive le mois précédent s'il ne l'est pas encore.
        /// Sans risque si appelé plusieurs fois ou depuis plusieurs postes.
        /// </summary>
        public static Task ArchiverMoisPrecedentSiNecessaireAsync()
        {
            return Task.Run(() =>
            {
                try
                {
                    var aujourd_hui = DateTime.Now;
                    var moisPrec = new DateTime(aujourd_hui.Year, aujourd_hui.Month, 1).AddMonths(-1);
                    ArchiverMoisVersDossierInterne(moisPrec);
                }
                catch (Exception ex)
                {
                    // Non bloquant : l'app reste utilisable si l'archivage échoue.
                    try
                    {
                        JournalLog.Warn(CategorieLog.Systeme, "ARCHIVE_LOGS_ERR",
                            $"Archivage mensuel échoué : {ex.Message}");
                    }
                    catch { /* le journal n'est peut-être pas encore configuré */ }
                }
            });
        }

        /// <summary>
        /// Déplace <c>entries-&lt;mois&gt;.jsonl</c> et <c>.txt</c> de <c>Logs/</c> vers
        /// <c>Logs/Archives/&lt;mois&gt;/</c>. Renvoie le nombre de fichiers déplacés
        /// (0 si déjà fait ou rien à déplacer).
        /// </summary>
        private static int ArchiverMoisVersDossierInterne(DateTime mois)
        {
            string moisStr = mois.ToString("yyyy-MM");
            string dossierLogs = CheminsMetrologo.Logs;
            string dossierArchive = Path.Combine(DossierArchivesRacine, moisStr);
            string flagDone = Path.Combine(dossierArchive, ".done");
            string flagLock = Path.Combine(dossierLogs, $"archive-{moisStr}.lock");

            if (!Directory.Exists(dossierLogs)) return 0;

            // Déjà archivé par un autre poste.
            if (File.Exists(flagDone)) return 0;

            var fichiersAArchiver = Directory.EnumerateFiles(dossierLogs, $"entries-{moisStr}.*").ToList();
            if (fichiersAArchiver.Count == 0) return 0;

            // Verrou exclusif : FileMode.CreateNew est atomique, un seul poste l'emporte.
            FileStream? lockStream = null;
            try
            {
                try
                {
                    lockStream = new FileStream(flagLock, FileMode.CreateNew,
                        FileAccess.Write, FileShare.None);
                }
                catch (IOException)
                {
                    // Un autre poste tient le verrou : c'est lui qui archive.
                    return 0;
                }

                // Revérification sous verrou : un poste a pu finir juste avant nous.
                if (File.Exists(flagDone)) return 0;

                Directory.CreateDirectory(dossierArchive);

                int nbDeplaces = 0;
                foreach (var fichier in fichiersAArchiver)
                {
                    string nom = Path.GetFileName(fichier);
                    string dest = Path.Combine(dossierArchive, nom);
                    try
                    {
                        if (File.Exists(dest)) File.Delete(dest);
                        File.Move(fichier, dest);
                        nbDeplaces++;
                    }
                    catch (Exception ex)
                    {
                        JournalLog.Warn(CategorieLog.Systeme, "ARCHIVE_LOGS_FICHIER_KO",
                            $"Impossible de déplacer {nom} vers archive : {ex.Message}");
                    }
                }

                // Marqueur de fin : les autres postes sauront que c'est fait.
                File.WriteAllText(flagDone,
                    $"Archivé le {DateTime.Now:yyyy-MM-dd HH:mm:ss} par {Environment.MachineName} "
                  + $"({nbDeplaces} fichier(s) déplacé(s))");

                JournalLog.Info(CategorieLog.Systeme, "ARCHIVE_LOGS_OK",
                    $"Logs de {moisStr} archivés ({nbDeplaces} fichier(s)) vers {dossierArchive}");

                return nbDeplaces;
            }
            finally
            {
                lockStream?.Dispose();
                try { if (File.Exists(flagLock)) File.Delete(flagLock); } catch { }
            }
        }


    }
}
