using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Dapper;
using Microsoft.Data.SqlClient;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Journal centralisé sur SQL Server — remplace <see cref="SqliteJournalService"/>
    /// pour permettre à l'admin de consulter l'activité de tous les postes (Baie,
    /// Paillasse, dev) depuis n'importe quelle machine.
    ///
    /// Convention horaire : tout est écrit en UTC, converti en heure locale à la
    /// lecture. La bascule été/hiver est ainsi gérée automatiquement par Windows.
    ///
    /// Robustesse : les écritures sont sérialisées via un sémaphore unique pour
    /// éviter les courses sur la session courante. En cas d'indisponibilité du
    /// serveur, les Log() fire-and-forget se loggent en console et continuent —
    /// la session courante reste active mais sans persistance pour les entrées
    /// concernées (TODO : cache local en cas de coupure si nécessaire).
    /// </summary>
    public class SqlServerJournalService : IJournalService
    {
        private readonly SemaphoreSlim _ecritureSema = new(1, 1);

        public string? SessionActuelleId { get; private set; }
        public string? UtilisateurActuel { get; private set; }

        public SqlServerJournalService()
        {
            // Le schéma est créé par DatabaseInitializer au démarrage de l'app —
            // ce ctor reste léger pour ne pas bloquer si SQL Server est temporairement
            // indisponible (fallback : sessions/logs ignorés silencieusement).
        }

        // -------------------------------------------------------------------------
        // Sessions
        // -------------------------------------------------------------------------

        public async Task NettoyerSessionsZombiesAsync()
        {
            await _ecritureSema.WaitAsync();
            try
            {
                using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
                await c.OpenAsync();
                // Pour chaque session sans Fin, on fixe Fin = max(Timestamp des entries)
                // ou Debut si la session n'a aucune entrée. Idempotent.
                await c.ExecuteAsync(@"
                    UPDATE s
                    SET Fin = COALESCE(
                        (SELECT MAX(e.Timestamp) FROM dbo.T_LOGS_ENTRIES e WHERE e.SessionId = s.SessionId),
                        s.Debut)
                    FROM dbo.T_LOGS_SESSIONS s
                    WHERE s.Fin IS NULL");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[Journal] NettoyerSessionsZombies KO : {ex.Message}");
            }
            finally { _ecritureSema.Release(); }
        }

        public async Task DemarrerSessionAsync(string utilisateur)
        {
            SessionActuelleId = Guid.NewGuid().ToString("N");
            UtilisateurActuel = utilisateur;

            await _ecritureSema.WaitAsync();
            try
            {
                using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
                await c.OpenAsync();
                await c.ExecuteAsync(
                    @"INSERT INTO dbo.T_LOGS_SESSIONS (SessionId, Utilisateur, Machine, Debut)
                      VALUES (@SessionId, @Utilisateur, @Machine, @Debut)",
                    new
                    {
                        SessionId = SessionActuelleId,
                        Utilisateur = utilisateur,
                        Machine = Environment.MachineName,
                        Debut = DateTime.UtcNow
                    });
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[Journal] DemarrerSession KO : {ex.Message}");
            }
            finally { _ecritureSema.Release(); }

            await LogAsync(CategorieLog.Session, "DEBUT_SESSION",
                $"Session ouverte pour {utilisateur} sur {Environment.MachineName}");
        }

        public async Task TerminerSessionAsync()
        {
            if (SessionActuelleId == null) return;

            await LogAsync(CategorieLog.Session, "FIN_SESSION", "Session fermée.");

            await _ecritureSema.WaitAsync();
            try
            {
                using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
                await c.OpenAsync();
                await c.ExecuteAsync(
                    "UPDATE dbo.T_LOGS_SESSIONS SET Fin = @Fin WHERE SessionId = @Id",
                    new { Fin = DateTime.UtcNow, Id = SessionActuelleId });
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[Journal] TerminerSession KO : {ex.Message}");
            }
            finally { _ecritureSema.Release(); }

            SessionActuelleId = null;
            UtilisateurActuel = null;
        }

        public async Task DefinirPosteAsync(string poste)
        {
            if (SessionActuelleId == null) return;
            await _ecritureSema.WaitAsync();
            try
            {
                using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
                await c.OpenAsync();
                await c.ExecuteAsync(
                    "UPDATE dbo.T_LOGS_SESSIONS SET Poste = @P WHERE SessionId = @Id",
                    new { P = poste, Id = SessionActuelleId });
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[Journal] DefinirPoste KO : {ex.Message}");
            }
            finally { _ecritureSema.Release(); }
        }

        // -------------------------------------------------------------------------
        // Log entries
        // -------------------------------------------------------------------------

        public void Log(CategorieLog c, string a, string m, object? d = null, SeveriteLog s = SeveriteLog.Info)
        {
            // Fire-and-forget : la console n'attend pas la BD. Les éventuelles erreurs
            // SQL sont loggées en console mais ne propagent pas.
            _ = LogAsync(c, a, m, d, s);
        }

        public async Task LogAsync(CategorieLog categorie, string action, string message,
            object? details = null, SeveriteLog severite = SeveriteLog.Info)
        {
            if (SessionActuelleId == null) return;

            string? detailsJson = null;
            if (details != null)
            {
                try { detailsJson = JsonSerializer.Serialize(details); }
                catch { detailsJson = details.ToString(); }
            }

            await _ecritureSema.WaitAsync();
            try
            {
                using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
                await c.OpenAsync();
                await c.ExecuteAsync(
                    @"INSERT INTO dbo.T_LOGS_ENTRIES
                      (SessionId, Timestamp, Categorie, Action, Message, Details, Severite)
                      VALUES (@SessionId, @Ts, @Cat, @Action, @Msg, @Det, @Sev)",
                    new
                    {
                        SessionId = SessionActuelleId,
                        Ts = DateTime.UtcNow,
                        Cat = categorie.ToString(),
                        Action = action,
                        Msg = message ?? "",
                        Det = (object?)detailsJson ?? DBNull.Value,
                        Sev = severite.ToString()
                    });
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[Journal] Log KO ({action}) : {ex.Message}");
            }
            finally { _ecritureSema.Release(); }
        }

        // -------------------------------------------------------------------------
        // Lecture / consultation
        // -------------------------------------------------------------------------

        public async Task<List<SessionJournal>> ChargerSessionsAsync(FiltreJournal? filtre = null)
        {
            filtre ??= new FiltreJournal();
            var sessions = new Dictionary<string, SessionJournal>();

            using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
            await c.OpenAsync();

            // ------------ Sessions ------------
            var whereSessions = "1 = 1";
            var paramsSessions = new DynamicParameters();
            if (filtre.Depuis.HasValue)
            {
                whereSessions += " AND Debut >= @Depuis";
                // Filtre venu de l'UI en heure locale → on convertit en UTC pour matcher.
                paramsSessions.Add("Depuis", filtre.Depuis.Value.ToUniversalTime());
            }
            if (filtre.Jusqu_a.HasValue)
            {
                whereSessions += " AND Debut <= @Jusqu";
                paramsSessions.Add("Jusqu", filtre.Jusqu_a.Value.ToUniversalTime());
            }
            if (!string.IsNullOrEmpty(filtre.Utilisateur))
            {
                whereSessions += " AND Utilisateur = @Utilisateur";
                paramsSessions.Add("Utilisateur", filtre.Utilisateur);
            }

            var sessionRows = await c.QueryAsync<SessionRow>(
                $@"SELECT TOP 200 SessionId, Utilisateur, Machine, Debut, Fin, Poste
                   FROM dbo.T_LOGS_SESSIONS
                   WHERE {whereSessions}
                   ORDER BY Debut DESC",
                paramsSessions);

            foreach (var r in sessionRows)
            {
                var s = new SessionJournal
                {
                    SessionId = r.SessionId,
                    Utilisateur = r.Utilisateur,
                    Machine = r.Machine,
                    Debut = ToLocal(r.Debut),
                    Fin = r.Fin.HasValue ? ToLocal(r.Fin.Value) : (DateTime?)null,
                    Poste = r.Poste
                };
                sessions[s.SessionId] = s;
            }
            if (sessions.Count == 0) return new List<SessionJournal>();

            // ------------ Entries ------------
            var whereEntries = "SessionId IN @SessionIds";
            var paramsEntries = new DynamicParameters();
            paramsEntries.Add("SessionIds", sessions.Keys.ToList());

            if (filtre.Categorie.HasValue)
            {
                whereEntries += " AND Categorie = @Cat";
                paramsEntries.Add("Cat", filtre.Categorie.Value.ToString());
            }
            if (filtre.SeveriteMin.HasValue)
            {
                var sevs = Enum.GetValues(typeof(SeveriteLog)).Cast<SeveriteLog>()
                    .Where(s => s >= filtre.SeveriteMin.Value).Select(s => s.ToString()).ToList();
                whereEntries += " AND Severite IN @Sevs";
                paramsEntries.Add("Sevs", sevs);
            }
            if (!string.IsNullOrEmpty(filtre.Recherche))
            {
                whereEntries += " AND (Message LIKE @R OR Action LIKE @R OR Details LIKE @R)";
                paramsEntries.Add("R", $"%{filtre.Recherche}%");
            }

            // Filtre métier : ActionsMetier (whitelist) OR Severite >= Avertissement.
            // Évite de matérialiser des milliers d'entrées techniques en mode normal.
            if (filtre.ActionsMetier != null && filtre.ActionsMetier.Count > 0)
            {
                whereEntries += " AND (Action IN @ActionsMetier OR Severite IN @SevsHaut)";
                paramsEntries.Add("ActionsMetier", filtre.ActionsMetier);
                paramsEntries.Add("SevsHaut", new[] { SeveriteLog.Avertissement.ToString(), SeveriteLog.Erreur.ToString() });
            }

            var entryRows = await c.QueryAsync<EntryRow>(
                $@"SELECT EntryId, SessionId, Timestamp, Categorie, Action, Message, Details, Severite
                   FROM dbo.T_LOGS_ENTRIES
                   WHERE {whereEntries}
                   ORDER BY Timestamp ASC",
                paramsEntries);

            foreach (var e in entryRows)
            {
                if (!sessions.TryGetValue(e.SessionId, out var s)) continue;
                s.Entrees.Add(new LogEntry
                {
                    EntryId = e.EntryId,
                    SessionId = e.SessionId,
                    Timestamp = ToLocal(e.Timestamp),
                    Categorie = Enum.TryParse<CategorieLog>(e.Categorie, out var cat) ? cat : CategorieLog.Systeme,
                    Action = e.Action,
                    Message = e.Message ?? "",
                    Details = e.Details,
                    Severite = Enum.TryParse<SeveriteLog>(e.Severite, out var sev) ? sev : SeveriteLog.Info
                });
            }

            var result = new List<SessionJournal>(sessions.Values);
            result.Sort((a, b) => b.Debut.CompareTo(a.Debut));
            return result;
        }

        public async Task<List<string>> ChargerListeUtilisateursAsync()
        {
            using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
            await c.OpenAsync();
            var users = await c.QueryAsync<string>(
                "SELECT DISTINCT Utilisateur FROM dbo.T_LOGS_SESSIONS ORDER BY Utilisateur");
            return users.ToList();
        }

        // -------------------------------------------------------------------------
        // Helpers
        // -------------------------------------------------------------------------

        /// <summary>
        /// Les colonnes DATETIME2 SQL Server sont remontées par Dapper en
        /// <see cref="DateTimeKind.Unspecified"/>. On marque Kind=Utc puis conversion
        /// en heure locale du poste — Windows gère DST automatiquement.
        /// </summary>
        private static DateTime ToLocal(DateTime utcUnspecified)
            => DateTime.SpecifyKind(utcUnspecified, DateTimeKind.Utc).ToLocalTime();

        // DTOs internes Dapper
        private sealed class SessionRow
        {
            public string SessionId { get; set; } = "";
            public string Utilisateur { get; set; } = "";
            public string Machine { get; set; } = "";
            public DateTime Debut { get; set; }
            public DateTime? Fin { get; set; }
            public string? Poste { get; set; }
        }

        private sealed class EntryRow
        {
            public long EntryId { get; set; }
            public string SessionId { get; set; } = "";
            public DateTime Timestamp { get; set; }
            public string Categorie { get; set; } = "";
            public string Action { get; set; } = "";
            public string? Message { get; set; }
            public string? Details { get; set; }
            public string Severite { get; set; } = "";
        }
    }
}
