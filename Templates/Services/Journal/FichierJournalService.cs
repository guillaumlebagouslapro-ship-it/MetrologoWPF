using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Metrologo.Services;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Journal centralisé sur fichiers JSON dans CheminsMetrologo.Logs (par défaut
    /// M:\exe_spe\Data_Metrologo\Logs\). Remplace l'ancien SqlServerJournalService.
    /// </summary>
    // Stockage : sessions.json (tableau complet relu/réécrit en bloc, une session = une
    // ouverture d'app sur un poste) + entries-YYYY-MM.jsonl (JSON-lines append-only, un
    // fichier par mois : plus rapide en écriture et archivable mois par mois).
    // Multi-postes : sessions identifiées par Guid (pas de collision), entries en append
    // avec FileShare.Read donc plusieurs postes écrivent en parallèle sans bloquer la lecture.
    // Si le partage M:\ est HS, les écritures échouent en silence et l'app reste utilisable.
    public class FichierJournalService : IJournalService
    {
        private readonly SemaphoreSlim _semaSessions = new(1, 1);

        public string? SessionActuelleId { get; private set; }
        public string? UtilisateurActuel { get; private set; }

        private static readonly JsonSerializerOptions _jsonOpts = new()
        {
            WriteIndented = false,
            PropertyNamingPolicy = null
        };

        // ---- Sessions ----

        public async Task DemarrerSessionAsync(string utilisateur)
        {
            SessionActuelleId = Guid.NewGuid().ToString("N");
            UtilisateurActuel = utilisateur;

            var session = new SessionJournal
            {
                SessionId = SessionActuelleId,
                Utilisateur = utilisateur,
                Machine = Environment.MachineName,
                Debut = DateTime.Now,
                Fin = null,
            };

            await ModifierSessionsAsync(sessions =>
            {
                sessions.Add(session);
            });
        }

        public async Task TerminerSessionAsync()
        {
            if (string.IsNullOrEmpty(SessionActuelleId)) return;
            string idCourant = SessionActuelleId;

            await ModifierSessionsAsync(sessions =>
            {
                var cible = sessions.FirstOrDefault(s => s.SessionId == idCourant);
                if (cible != null && cible.Fin == null)
                {
                    cible.Fin = DateTime.Now;
                }
            });

            SessionActuelleId = null;
            UtilisateurActuel = null;
        }

        public async Task DefinirPosteAsync(string poste)
        {
            if (string.IsNullOrEmpty(SessionActuelleId)) return;
            string idCourant = SessionActuelleId;

            await ModifierSessionsAsync(sessions =>
            {
                var cible = sessions.FirstOrDefault(s => s.SessionId == idCourant);
                if (cible != null) cible.Poste = poste;
            });
        }

        public async Task NettoyerSessionsZombiesAsync()
        {
            // Deux cas :
            // - cette machine : si on est la seule instance Metrologo qui tourne, toute
            //   session locale sans Fin est forcément zombie (fermeture brutale).
            // - autres machines : impossible de savoir si l'app y tourne encore, donc
            //   seuil généreux de 24h avant de fermer.
            // Pour chaque zombie, Fin = dernier Timestamp d'entrée, ou Debut si aucune entrée.

            string maMachine = Environment.MachineName;
            int monPid = System.Diagnostics.Process.GetCurrentProcess().Id;
            bool seuleInstanceLocale = System.Diagnostics.Process
                .GetProcessesByName("Metrologo")
                .All(p => p.Id == monPid);

            var seuilZombieAutresMachines = DateTime.Now.AddHours(-24);
            var dernierTimestampParSession = await CalculerDerniersTimestampsAsync();

            await ModifierSessionsAsync(sessions =>
            {
                foreach (var s in sessions)
                {
                    if (s.Fin != null) continue;

                    bool estCetteMachine = string.Equals(s.Machine, maMachine,
                        StringComparison.OrdinalIgnoreCase);

                    bool aFermer;
                    if (estCetteMachine && seuleInstanceLocale)
                    {
                        // On vient de démarrer la seule instance sur cette machine →
                        // toute session « en cours » d'ici est forcément zombie.
                        aFermer = true;
                    }
                    else
                    {
                        // Autres machines (ou cette machine mais autres instances actives) :
                        // on attend le seuil 24h pour ne pas fermer prématurément.
                        aFermer = s.Debut < seuilZombieAutresMachines;
                    }

                    if (aFermer)
                    {
                        s.Fin = dernierTimestampParSession.TryGetValue(s.SessionId, out var t)
                            ? t
                            : s.Debut;
                    }
                }
            });
        }

        // -------------------------------------------------------------------------
        // Entries
        // -------------------------------------------------------------------------

        public Task LogAsync(CategorieLog categorie, string action, string message,
                              object? details = null, SeveriteLog severite = SeveriteLog.Info)
            => EcrireEntreeAsync(categorie, action, message, details, severite);

        public void Log(CategorieLog categorie, string action, string message,
                         object? details = null, SeveriteLog severite = SeveriteLog.Info)
        {
            // Fire-and-forget — la perte d'une entrée n'est pas critique.
            _ = EcrireEntreeAsync(categorie, action, message, details, severite);
        }

        private Task EcrireEntreeAsync(CategorieLog categorie, string action, string message,
                                       object? details, SeveriteLog severite)
        {
            return Task.Run(() =>
            {
                try
                {
                    var entry = new LogEntry
                    {
                        EntryId = DateTime.UtcNow.Ticks,
                        SessionId = SessionActuelleId ?? string.Empty,
                        Timestamp = DateTime.Now,
                        Categorie = categorie,
                        Action = action,
                        Message = message ?? string.Empty,
                        Details = details == null ? null : JsonSerializer.Serialize(details, _jsonOpts),
                        Severite = severite,
                    };

                    string fichier = FichierEntreesDuMois(entry.Timestamp);
                    string fichierTxt = FichierEntreesTexteDuMois(entry.Timestamp);
                    string? dossier = Path.GetDirectoryName(fichier);
                    if (!string.IsNullOrEmpty(dossier)) Directory.CreateDirectory(dossier);

                    // 1) JSON-lines (machine-readable, consommé par le JournalViewer)
                    string ligne = JsonSerializer.Serialize(entry, _jsonOpts);
                    // Append concurrent-safe : FileShare.Read autorise les lectures simultanées,
                    // pas d'écriture concurrente (un seul writer à la fois côté Windows).
                    using (var fs = new FileStream(fichier, FileMode.Append, FileAccess.Write,
                        FileShare.Read, bufferSize: 4096, useAsync: false))
                    using (var sw = new StreamWriter(fs, Encoding.UTF8))
                    {
                        sw.WriteLine(ligne);
                    }

                    // 2) TXT human-readable (consultable directement avec Notepad / tail / grep)
                    string ligneTxt = FormaterLigneTexte(entry);
                    try
                    {
                        using var fsTxt = new FileStream(fichierTxt, FileMode.Append, FileAccess.Write,
                            FileShare.Read, bufferSize: 4096, useAsync: false);
                        using var swTxt = new StreamWriter(fsTxt, Encoding.UTF8);
                        swTxt.WriteLine(ligneTxt);
                    }
                    catch { /* best-effort : la perte du TXT n'affecte pas le JSON */ }
                }
                catch
                {
                    // Partage HS / fichier verrouillé → on perd cette entrée. L'app continue.
                }
            });
        }

        private static string FichierEntreesDuMois(DateTime moment) =>
            Path.Combine(CheminsMetrologo.Logs, $"entries-{moment:yyyy-MM}.jsonl");

        /// <summary>
        /// Chemin du fichier TXT lisible accompagnant le JSON-lines du même mois.
        /// Permet de consulter les logs directement avec Notepad / PowerShell sans
        /// passer par l'application Metrologo.
        /// </summary>
        private static string FichierEntreesTexteDuMois(DateTime moment) =>
            Path.Combine(CheminsMetrologo.Logs, $"entries-{moment:yyyy-MM}.txt");

        /// <summary>
        /// Formate une entrée pour le fichier TXT lisible. Pattern style log4j :
        /// <c>2026-05-25 14:32:47.123 [INFO ] [Administration ] [ACCES_ADMIN_OK     ] message · session=… machine=…</c>
        /// Champs alignés en largeur fixe pour faciliter le grep / tri par colonne.
        /// </summary>
        private string FormaterLigneTexte(LogEntry e)
        {
            string ts  = e.Timestamp.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string sev = e.Severite switch
            {
                SeveriteLog.Erreur        => "ERROR",
                SeveriteLog.Avertissement => "WARN ",
                _                          => "INFO ",
            };
            string cat = e.Categorie.ToString().PadRight(16);
            string act = (e.Action ?? string.Empty).PadRight(28);
            string msg = (e.Message ?? string.Empty).Replace("\r", " ").Replace("\n", " ");

            var sb = new StringBuilder();
            sb.Append(ts).Append(' ')
              .Append('[').Append(sev).Append(']').Append(' ')
              .Append('[').Append(cat).Append(']').Append(' ')
              .Append('[').Append(act).Append(']').Append(' ')
              .Append(msg);

            // Metadata utile pour corrélation multi-postes
            sb.Append("  ·  user=").Append(UtilisateurActuel ?? "?");
            sb.Append(" machine=").Append(Environment.MachineName);
            if (!string.IsNullOrEmpty(e.SessionId))
                sb.Append(" session=").Append(e.SessionId.Substring(0, Math.Min(8, e.SessionId.Length)));

            // Détails JSON inline si court ; sinon résumé
            if (!string.IsNullOrEmpty(e.Details))
            {
                string d = e.Details.Replace("\r", " ").Replace("\n", " ");
                if (d.Length > 200) d = d.Substring(0, 200) + "…";
                sb.Append(" details=").Append(d);
            }
            return sb.ToString();
        }

        // -------------------------------------------------------------------------
        // Lecture pour JournalViewer
        // -------------------------------------------------------------------------

        public async Task<List<SessionJournal>> ChargerSessionsAsync(FiltreJournal? filtre = null)
        {
            await _semaSessions.WaitAsync();
            try
            {
                var sessions = LireFichierSessions();

                // Filtre temporel + utilisateur sur les sessions
                if (filtre != null)
                {
                    if (filtre.Depuis.HasValue)
                        sessions = sessions.Where(s => s.Debut >= filtre.Depuis.Value).ToList();
                    if (filtre.Jusqu_a.HasValue)
                        sessions = sessions.Where(s => s.Debut <= filtre.Jusqu_a.Value).ToList();
                    if (!string.IsNullOrWhiteSpace(filtre.Utilisateur))
                        sessions = sessions.Where(s =>
                            string.Equals(s.Utilisateur, filtre.Utilisateur, StringComparison.OrdinalIgnoreCase))
                            .ToList();
                }

                // Charge les entrées des mois pertinents
                var sessionsParId = sessions.ToDictionary(s => s.SessionId);
                foreach (var entry in LireToutesLesEntrees(filtre))
                {
                    if (sessionsParId.TryGetValue(entry.SessionId, out var s))
                    {
                        s.Entrees.Add(entry);
                    }
                }

                return sessions.OrderByDescending(s => s.Debut).ToList();
            }
            finally
            {
                _semaSessions.Release();
            }
        }

        public Task<List<string>> ChargerListeUtilisateursAsync()
        {
            var sessions = LireFichierSessions();
            var users = sessions
                .Select(s => s.Utilisateur)
                .Where(u => !string.IsNullOrWhiteSpace(u))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(u => u)
                .ToList();
            return Task.FromResult(users);
        }

        // -------------------------------------------------------------------------
        // I/O sessions.json
        // -------------------------------------------------------------------------

        private static string FichierSessions =>
            Path.Combine(CheminsMetrologo.Logs, "sessions.json");

        private List<SessionJournal> LireFichierSessions()
        {
            try
            {
                string chemin = FichierSessions;
                if (!File.Exists(chemin)) return new List<SessionJournal>();
                string json = File.ReadAllText(chemin);
                if (string.IsNullOrWhiteSpace(json)) return new List<SessionJournal>();
                return JsonSerializer.Deserialize<List<SessionJournal>>(json, _jsonOpts)
                       ?? new List<SessionJournal>();
            }
            catch
            {
                return new List<SessionJournal>();
            }
        }

        /// <summary>
        /// Lit / modifie / réécrit le fichier sessions.json de manière atomique.
        /// Le sémaphore garantit qu'un seul thread du même poste y touche à la fois.
        /// </summary>
        private async Task ModifierSessionsAsync(Action<List<SessionJournal>> modification)
        {
            await _semaSessions.WaitAsync();
            try
            {
                var sessions = LireFichierSessions();
                modification(sessions);

                string chemin = FichierSessions;
                string? dossier = Path.GetDirectoryName(chemin);
                if (!string.IsNullOrEmpty(dossier)) Directory.CreateDirectory(dossier);

                string tmp = chemin + ".tmp";
                string json = JsonSerializer.Serialize(sessions, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = null,
                });
                await File.WriteAllTextAsync(tmp, json);

                if (File.Exists(chemin))
                {
                    try { File.Replace(tmp, chemin, destinationBackupFileName: null); }
                    catch (PlatformNotSupportedException)
                    {
                        File.Delete(chemin);
                        File.Move(tmp, chemin);
                    }
                }
                else
                {
                    File.Move(tmp, chemin);
                }
            }
            catch
            {
                // Best-effort : partage HS / fichier verrouillé → l'opération est perdue.
            }
            finally
            {
                _semaSessions.Release();
            }
        }

        // -------------------------------------------------------------------------
        // I/O entries-YYYY-MM.jsonl
        // -------------------------------------------------------------------------

        private IEnumerable<LogEntry> LireToutesLesEntrees(FiltreJournal? filtre)
        {
            string dossier = CheminsMetrologo.Logs;
            if (!Directory.Exists(dossier)) yield break;

            // On collecte tous les fichiers entries-*.jsonl à lire :
            //   1) Mois actifs/récents : directement dans Logs/
            //   2) Mois archivés : dans Logs/Archives/YYYY-MM/
            // Comme ça le viewer voit l'historique complet, même après rotation mensuelle.
            var fichiersAlire = new List<string>(Directory.EnumerateFiles(dossier, "entries-*.jsonl"));
            string dossierArchives = Path.Combine(dossier, "Archives");
            if (Directory.Exists(dossierArchives))
            {
                foreach (var sd in Directory.EnumerateDirectories(dossierArchives))
                {
                    fichiersAlire.AddRange(Directory.EnumerateFiles(sd, "entries-*.jsonl"));
                }
            }

            // Filtres temporels appliqués au niveau du nom de fichier pour zapper les mois
            // non pertinents avant de lire les lignes.
            DateTime? bornInf = filtre?.Depuis;
            DateTime? bornSup = filtre?.Jusqu_a;

            foreach (var fichier in fichiersAlire)
            {
                // Optim : filtre par nom de fichier sur le mois si possible
                if (bornInf.HasValue || bornSup.HasValue)
                {
                    if (TryExtraireMois(Path.GetFileName(fichier), out var mois))
                    {
                        var debutMois = new DateTime(mois.Year, mois.Month, 1);
                        var finMois = debutMois.AddMonths(1).AddTicks(-1);
                        if (bornInf.HasValue && finMois < bornInf.Value) continue;
                        if (bornSup.HasValue && debutMois > bornSup.Value) continue;
                    }
                }

                IEnumerable<string> lignes;
                try { lignes = File.ReadLines(fichier); }
                catch { continue; }

                foreach (var ligne in lignes)
                {
                    if (string.IsNullOrWhiteSpace(ligne)) continue;
                    LogEntry? entry;
                    try { entry = JsonSerializer.Deserialize<LogEntry>(ligne, _jsonOpts); }
                    catch { continue; }
                    if (entry == null) continue;

                    if (filtre != null && !PasseLesFiltres(entry, filtre)) continue;
                    yield return entry;
                }
            }
        }

        private static bool TryExtraireMois(string nomFichier, out DateTime mois)
        {
            // entries-YYYY-MM.jsonl
            mois = default;
            const string prefix = "entries-";
            const string suffix = ".jsonl";
            if (!nomFichier.StartsWith(prefix) || !nomFichier.EndsWith(suffix)) return false;
            string moisStr = nomFichier.Substring(prefix.Length, nomFichier.Length - prefix.Length - suffix.Length);
            return DateTime.TryParseExact(moisStr, "yyyy-MM",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out mois);
        }

        private static bool PasseLesFiltres(LogEntry e, FiltreJournal f)
        {
            if (f.Depuis.HasValue && e.Timestamp < f.Depuis.Value) return false;
            if (f.Jusqu_a.HasValue && e.Timestamp > f.Jusqu_a.Value) return false;
            if (f.Categorie.HasValue && e.Categorie != f.Categorie.Value) return false;
            if (f.SeveriteMin.HasValue && (int)e.Severite < (int)f.SeveriteMin.Value) return false;
            if (!string.IsNullOrWhiteSpace(f.Recherche))
            {
                var r = f.Recherche;
                if (!(e.Message?.Contains(r, StringComparison.OrdinalIgnoreCase) == true
                   || e.Action?.Contains(r, StringComparison.OrdinalIgnoreCase) == true
                   || e.Details?.Contains(r, StringComparison.OrdinalIgnoreCase) == true))
                    return false;
            }
            if (f.ActionsMetier != null
                && !f.ActionsMetier.Contains(e.Action, StringComparer.OrdinalIgnoreCase)
                && e.Severite < SeveriteLog.Avertissement)
            {
                return false;
            }
            return true;
        }

        private async Task<Dictionary<string, DateTime>> CalculerDerniersTimestampsAsync()
        {
            // Pour chaque session, on garde le dernier Timestamp d'entrée — utile
            // pour fermer proprement les sessions zombies.
            var result = new Dictionary<string, DateTime>();
            await Task.Run(() =>
            {
                foreach (var e in LireToutesLesEntrees(null))
                {
                    if (string.IsNullOrEmpty(e.SessionId)) continue;
                    if (!result.TryGetValue(e.SessionId, out var existant) || e.Timestamp > existant)
                        result[e.SessionId] = e.Timestamp;
                }
            });
            return result;
        }
    }
}
