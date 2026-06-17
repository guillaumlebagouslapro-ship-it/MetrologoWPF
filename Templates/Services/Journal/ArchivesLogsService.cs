using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Range les fichiers de logs du mois écoulé dans un sous-dossier d'archive, sur le partage réseau.
    ///
    /// L'idée : au tout premier lancement de Metrologo dans le mois (n'importe quel poste,
    /// n'importe quel utilisateur), on déplace les fichiers <c>entries-&lt;moisPrec&gt;.jsonl</c>
    /// et <c>.txt</c> de <c>M:\Logs\</c> vers <c>M:\Logs\Archives\&lt;moisPrec&gt;\</c>. Les logs
    /// du mois en cours, eux, continuent de s'écrire dans <c>M:\Logs\</c> comme d'habitude.
    ///
    /// Pensé pour le multi-postes : un verrou fichier exclusif (<c>archive-YYYY-MM.lock</c>) et
    /// un marqueur de fin (<c>Archives\YYYY-MM\.done</c>). Si trois postes démarrent le matin du
    /// 1er, un seul fait le travail ; les autres tombent sur le .done et passent leur chemin.
    /// Et si M:\ n'est pas accessible, on ne fait rien, sans bruit.
    /// </summary>
    public static class ArchivesLogsService
    {
        /// <summary>Le dossier Archives, situé à l'intérieur de Logs (sur le partage).</summary>
        public static string DossierArchivesRacine => Path.Combine(CheminsMetrologo.Logs, "Archives");

        // ---------- API publique, gardée telle quelle pour les ViewModels existants ----------

        /// <summary>
        /// Énumère les mois qui ont un fichier d'entrées : le mois actif (dans <c>Logs/</c>)
        /// et les mois déjà archivés (dans <c>Logs/Archives/YYYY-MM/</c>).
        /// </summary>
        public static IEnumerable<DateTime> ListerMoisArchives()
        {
            var dejaVus = new HashSet<DateTime>();
            string dossierLogs = CheminsMetrologo.Logs;

            // Mois encore actifs (les entries-YYYY-MM.jsonl posés dans Logs/)
            if (Directory.Exists(dossierLogs))
            {
                foreach (var fichier in Directory.EnumerateFiles(dossierLogs, "entries-*.jsonl"))
                {
                    if (TryExtraireMoisDuNom(fichier, out var mois) && dejaVus.Add(mois))
                        yield return mois;
                }
            }

            // Mois déjà archivés (un sous-dossier Logs/Archives/YYYY-MM/ par mois)
            string dossierArchives = DossierArchivesRacine;
            if (Directory.Exists(dossierArchives))
            {
                foreach (var sd in Directory.EnumerateDirectories(dossierArchives))
                {
                    string nom = Path.GetFileName(sd);
                    if (DateTime.TryParseExact(nom, "yyyy-MM",
                            System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.None, out var mois)
                        && dejaVus.Add(mois))
                    {
                        yield return mois;
                    }
                }
            }
        }

        public static IEnumerable<DateTime> ListerJoursDuMoisArchive(DateTime mois)
            => Enumerable.Empty<DateTime>();

        public static Task<List<SessionJournal>> ChargerJournauxArchivesAsync(DateTime jour)
            => Task.FromResult(new List<SessionJournal>());

        public static Task<int> ArchiverMoisAsync(DateTime mois, bool force = false)
        {
            int n = ArchiverMoisVersDossierInterne(mois);
            return Task.FromResult(n);
        }

        /// <summary>
        /// À appeler au démarrage de l'app : archive le mois précédent s'il ne l'a pas
        /// encore été. On peut l'appeler sans risque plusieurs fois, et depuis plusieurs postes.
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
                    // On n'en fait pas un drame : même si l'archivage rate, l'app reste utilisable.
                    try
                    {
                        JournalLog.Warn(CategorieLog.Systeme, "ARCHIVE_LOGS_ERR",
                            $"Archivage mensuel échoué : {ex.Message}");
                    }
                    catch { /* le journal n'est peut-être pas encore configuré */ }
                }
            });
        }

        // ---------- Implémentation interne ----------

        /// <summary>
        /// Déplace les fichiers <c>entries-&lt;mois&gt;.jsonl</c> et <c>.txt</c> de
        /// <c>Logs/</c> vers <c>Logs/Archives/&lt;mois&gt;/</c>. Sûr en multi-postes grâce au
        /// verrou fichier exclusif. Renvoie le nombre de fichiers déplacés (0 si c'était
        /// déjà fait ou s'il n'y avait rien à bouger).
        /// </summary>
        private static int ArchiverMoisVersDossierInterne(DateTime mois)
        {
            string moisStr = mois.ToString("yyyy-MM");
            string dossierLogs = CheminsMetrologo.Logs;
            string dossierArchive = Path.Combine(DossierArchivesRacine, moisStr);
            string flagDone = Path.Combine(dossierArchive, ".done");
            string flagLock = Path.Combine(dossierLogs, $"archive-{moisStr}.lock");

            // 0. Garde-fou de base
            if (!Directory.Exists(dossierLogs)) return 0;

            // 1. Un autre poste a déjà fait l'archivage ? On laisse tomber.
            if (File.Exists(flagDone)) return 0;

            // 2. Rien à archiver pour ce mois ? On laisse tomber aussi.
            var fichiersAArchiver = Directory.EnumerateFiles(dossierLogs, $"entries-{moisStr}.*").ToList();
            if (fichiersAArchiver.Count == 0) return 0;

            // 3. Verrou exclusif avec FileMode.CreateNew (la création est atomique).
            //    Si deux postes tentent leur chance en même temps, un seul l'emporte.
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
                    // Un autre poste tient déjà le verrou — c'est lui qui archivera, on passe.
                    return 0;
                }

                // 4. On revérifie une fois le verrou tenu (un poste a pu finir juste avant nous)
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

                // 5. On pose le marqueur de fin (comme ça les autres postes sauront que c'est réglé)
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

        // ---------- Helpers ----------

        private static bool TryExtraireMoisDuNom(string cheminFichier, out DateTime mois)
        {
            mois = default;
            string nom = Path.GetFileNameWithoutExtension(cheminFichier);
            const string prefix = "entries-";
            if (!nom.StartsWith(prefix)) return false;
            string moisStr = nom.Substring(prefix.Length);
            return DateTime.TryParseExact(moisStr, "yyyy-MM",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out mois);
        }
    }
}
