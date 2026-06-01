using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Archivage mensuel des fichiers de logs sur le partage réseau.
    ///
    /// Comportement : au 1er démarrage de Metrologo dans le mois courant (peu importe
    /// quel poste, lequel utilisateur), les fichiers <c>entries-&lt;moisPrec&gt;.jsonl</c> et
    /// <c>.txt</c> sont déplacés depuis <c>M:\Logs\</c> vers
    /// <c>M:\Logs\Archives\&lt;moisPrec&gt;\</c>. Les fichiers du mois courant continuent
    /// de s'écrire dans <c>M:\Logs\</c> normalement.
    ///
    /// Multi-postes safe : verrou fichier exclusif (<c>archive-YYYY-MM.lock</c>) +
    /// marqueur de complétion (<c>Archives\YYYY-MM\.done</c>). Si 3 postes démarrent
    /// le matin du 1er, un seul fait l'archivage, les autres voient le .done et passent.
    /// Si M:\ est inaccessible, opération no-op silencieuse.
    /// </summary>
    public static class ArchivesLogsService
    {
        /// <summary>Pointe sur le dossier Archives à l'intérieur de Logs (sur le partage).</summary>
        public static string DossierArchivesRacine => Path.Combine(CheminsMetrologo.Logs, "Archives");

        // ---------- API publique conservée pour les ViewModels existants ----------

        /// <summary>
        /// Liste les mois pour lesquels un fichier d'entrées existe (mois actif dans
        /// <c>Logs/</c> + mois archivés dans <c>Logs/Archives/YYYY-MM/</c>).
        /// </summary>
        public static IEnumerable<DateTime> ListerMoisArchives()
        {
            var dejaVus = new HashSet<DateTime>();
            string dossierLogs = CheminsMetrologo.Logs;

            // Mois actifs (entries-YYYY-MM.jsonl dans Logs/)
            if (Directory.Exists(dossierLogs))
            {
                foreach (var fichier in Directory.EnumerateFiles(dossierLogs, "entries-*.jsonl"))
                {
                    if (TryExtraireMoisDuNom(fichier, out var mois) && dejaVus.Add(mois))
                        yield return mois;
                }
            }

            // Mois archivés (sous-dossiers Logs/Archives/YYYY-MM/)
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
        /// Appelé au démarrage de l'app. Archive le mois précédent s'il ne l'est pas
        /// encore. Idempotent + multi-postes-safe.
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
                    // Erreur non bloquante — l'app reste utilisable même si l'archivage échoue.
                    try
                    {
                        JournalLog.Warn(CategorieLog.Systeme, "ARCHIVE_LOGS_ERR",
                            $"Archivage mensuel échoué : {ex.Message}");
                    }
                    catch { /* journal pas encore configuré ? */ }
                }
            });
        }

        // ---------- Implémentation interne ----------

        /// <summary>
        /// Déplace les fichiers <c>entries-&lt;mois&gt;.jsonl</c> et <c>.txt</c> depuis
        /// <c>Logs/</c> vers <c>Logs/Archives/&lt;mois&gt;/</c>. Multi-postes-safe via
        /// verrou fichier exclusif. Retourne le nombre de fichiers déplacés (0 si
        /// déjà fait ou rien à déplacer).
        /// </summary>
        private static int ArchiverMoisVersDossierInterne(DateTime mois)
        {
            string moisStr = mois.ToString("yyyy-MM");
            string dossierLogs = CheminsMetrologo.Logs;
            string dossierArchive = Path.Combine(DossierArchivesRacine, moisStr);
            string flagDone = Path.Combine(dossierArchive, ".done");
            string flagLock = Path.Combine(dossierLogs, $"archive-{moisStr}.lock");

            // 0. Vérif basique
            if (!Directory.Exists(dossierLogs)) return 0;

            // 1. Déjà archivé par un autre poste ? → on passe
            if (File.Exists(flagDone)) return 0;

            // 2. Rien à archiver pour ce mois ? → on passe
            var fichiersAArchiver = Directory.EnumerateFiles(dossierLogs, $"entries-{moisStr}.*").ToList();
            if (fichiersAArchiver.Count == 0) return 0;

            // 3. Verrou exclusif via FileMode.CreateNew (atomic à la création).
            //    Si deux postes tentent en parallèle, un seul gagne.
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
                    // Un autre poste tient le lock — il fera l'archivage, on passe.
                    return 0;
                }

                // 4. Re-vérif sous lock (un autre poste pourrait avoir fini juste avant nous)
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

                // 5. Marqueur de complétion (les autres postes verront que c'est fait)
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
