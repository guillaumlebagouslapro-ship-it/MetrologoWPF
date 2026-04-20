using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Metrologo.Services.Journal
{
    public class SqliteJournalService : IJournalService
    {
        private readonly string _connectionString;
        private readonly SemaphoreSlim _ecritureSema = new(1, 1);

        public string? SessionActuelleId { get; private set; }
        public string? UtilisateurActuel { get; private set; }

        public SqliteJournalService()
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Metrologo");
            Directory.CreateDirectory(dir);
            var chemin = Path.Combine(dir, "journal.db");
            _connectionString = $"Data Source={chemin};Cache=Shared";

            InitialiserSchemaAsync().GetAwaiter().GetResult();
        }

        private async Task InitialiserSchemaAsync()
        {
            using var c = new SqliteConnection(_connectionString);
            await c.OpenAsync();

            using (var pragma = c.CreateCommand())
            {
                // WAL autorise plusieurs lecteurs simultanés pendant l'écriture
                pragma.CommandText = "PRAGMA journal_mode=WAL;";
                await pragma.ExecuteNonQueryAsync();
            }

            using var cmd = c.CreateCommand();
            cmd.CommandText = @"
                CREATE TABLE IF NOT EXISTS Sessions (
                    SessionId   TEXT PRIMARY KEY,
                    Utilisateur TEXT NOT NULL,
                    Machine     TEXT NOT NULL,
                    Debut       TEXT NOT NULL,
                    Fin         TEXT
                );
                CREATE TABLE IF NOT EXISTS LogEntries (
                    EntryId    INTEGER PRIMARY KEY AUTOINCREMENT,
                    SessionId  TEXT NOT NULL,
                    Timestamp  TEXT NOT NULL,
                    Categorie  TEXT NOT NULL,
                    Action     TEXT NOT NULL,
                    Message    TEXT,
                    Details    TEXT,
                    Severite   TEXT NOT NULL,
                    FOREIGN KEY (SessionId) REFERENCES Sessions(SessionId)
                );
                CREATE INDEX IF NOT EXISTS idx_logs_session   ON LogEntries(SessionId);
                CREATE INDEX IF NOT EXISTS idx_logs_timestamp ON LogEntries(Timestamp);
                CREATE INDEX IF NOT EXISTS idx_sessions_user  ON Sessions(Utilisateur);
            ";
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task DemarrerSessionAsync(string utilisateur)
        {
            SessionActuelleId = Guid.NewGuid().ToString("N");
            UtilisateurActuel = utilisateur;

            await _ecritureSema.WaitAsync();
            try
            {
                using var c = new SqliteConnection(_connectionString);
                await c.OpenAsync();
                using var cmd = c.CreateCommand();
                cmd.CommandText = "INSERT INTO Sessions (SessionId, Utilisateur, Machine, Debut) VALUES (@id, @u, @m, @d);";
                cmd.Parameters.AddWithValue("@id", SessionActuelleId);
                cmd.Parameters.AddWithValue("@u", utilisateur);
                cmd.Parameters.AddWithValue("@m", Environment.MachineName);
                cmd.Parameters.AddWithValue("@d", DateTime.Now.ToString("o"));
                await cmd.ExecuteNonQueryAsync();
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
                using var c = new SqliteConnection(_connectionString);
                await c.OpenAsync();
                using var cmd = c.CreateCommand();
                cmd.CommandText = "UPDATE Sessions SET Fin = @f WHERE SessionId = @id;";
                cmd.Parameters.AddWithValue("@f", DateTime.Now.ToString("o"));
                cmd.Parameters.AddWithValue("@id", SessionActuelleId);
                await cmd.ExecuteNonQueryAsync();
            }
            finally { _ecritureSema.Release(); }

            SessionActuelleId = null;
            UtilisateurActuel = null;
        }

        public void Log(CategorieLog c, string a, string m, object? d = null, SeveriteLog s = SeveriteLog.Info)
        {
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
                using var c = new SqliteConnection(_connectionString);
                await c.OpenAsync();
                using var cmd = c.CreateCommand();
                cmd.CommandText = @"INSERT INTO LogEntries (SessionId, Timestamp, Categorie, Action, Message, Details, Severite)
                                    VALUES (@s, @t, @c, @a, @m, @d, @sev);";
                cmd.Parameters.AddWithValue("@s", SessionActuelleId);
                cmd.Parameters.AddWithValue("@t", DateTime.Now.ToString("o"));
                cmd.Parameters.AddWithValue("@c", categorie.ToString());
                cmd.Parameters.AddWithValue("@a", action);
                cmd.Parameters.AddWithValue("@m", message ?? "");
                cmd.Parameters.AddWithValue("@d", (object?)detailsJson ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@sev", severite.ToString());
                await cmd.ExecuteNonQueryAsync();
            }
            finally { _ecritureSema.Release(); }
        }

        public async Task<List<SessionJournal>> ChargerSessionsAsync(FiltreJournal? filtre = null)
        {
            filtre ??= new FiltreJournal();
            var sessions = new Dictionary<string, SessionJournal>();

            using var c = new SqliteConnection(_connectionString);
            await c.OpenAsync();

            // Charger les sessions
            using (var cmd = c.CreateCommand())
            {
                var where = "1=1";
                if (filtre.Depuis.HasValue) where += " AND Debut >= @depuis";
                if (filtre.Jusqu_a.HasValue) where += " AND Debut <= @jusqu";
                if (!string.IsNullOrEmpty(filtre.Utilisateur)) where += " AND Utilisateur = @u";

                cmd.CommandText = $"SELECT SessionId, Utilisateur, Machine, Debut, Fin FROM Sessions WHERE {where} ORDER BY Debut DESC LIMIT 200;";
                if (filtre.Depuis.HasValue) cmd.Parameters.AddWithValue("@depuis", filtre.Depuis.Value.ToString("o"));
                if (filtre.Jusqu_a.HasValue) cmd.Parameters.AddWithValue("@jusqu", filtre.Jusqu_a.Value.ToString("o"));
                if (!string.IsNullOrEmpty(filtre.Utilisateur)) cmd.Parameters.AddWithValue("@u", filtre.Utilisateur);

                using var reader = await cmd.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    var s = new SessionJournal
                    {
                        SessionId = reader.GetString(0),
                        Utilisateur = reader.GetString(1),
                        Machine = reader.GetString(2),
                        Debut = DateTime.Parse(reader.GetString(3)),
                        Fin = reader.IsDBNull(4) ? null : DateTime.Parse(reader.GetString(4))
                    };
                    sessions[s.SessionId] = s;
                }
            }

            if (sessions.Count == 0) return new List<SessionJournal>();

            // Charger les entrées
            using (var cmd = c.CreateCommand())
            {
                var ids = string.Join(",", sessions.Keys);
                // Param binding pour les IN() n'étant pas simple en SQLite, on utilise une chaîne (sessions déjà issues d'un WHERE)
                var placeholders = new List<string>();
                int i = 0;
                foreach (var k in sessions.Keys)
                {
                    var p = $"@p{i++}";
                    placeholders.Add(p);
                    cmd.Parameters.AddWithValue(p, k);
                }

                var where = $"SessionId IN ({string.Join(",", placeholders)})";
                if (filtre.Categorie.HasValue)
                {
                    where += " AND Categorie = @cat";
                    cmd.Parameters.AddWithValue("@cat", filtre.Categorie.Value.ToString());
                }
                if (filtre.SeveriteMin.HasValue)
                {
                    var sevs = new List<string>();
                    foreach (SeveriteLog s in Enum.GetValues(typeof(SeveriteLog)))
                        if (s >= filtre.SeveriteMin.Value) sevs.Add($"'{s}'");
                    where += $" AND Severite IN ({string.Join(",", sevs)})";
                }
                if (!string.IsNullOrEmpty(filtre.Recherche))
                {
                    where += " AND (Message LIKE @r OR Action LIKE @r OR Details LIKE @r)";
                    cmd.Parameters.AddWithValue("@r", $"%{filtre.Recherche}%");
                }

                cmd.CommandText = $"SELECT EntryId, SessionId, Timestamp, Categorie, Action, Message, Details, Severite FROM LogEntries WHERE {where} ORDER BY Timestamp ASC;";

                using var reader = await cmd.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    var sid = reader.GetString(1);
                    if (!sessions.TryGetValue(sid, out var s)) continue;

                    s.Entrees.Add(new LogEntry
                    {
                        EntryId = reader.GetInt64(0),
                        SessionId = sid,
                        Timestamp = DateTime.Parse(reader.GetString(2)),
                        Categorie = Enum.TryParse<CategorieLog>(reader.GetString(3), out var cat) ? cat : CategorieLog.Systeme,
                        Action = reader.GetString(4),
                        Message = reader.IsDBNull(5) ? string.Empty : reader.GetString(5),
                        Details = reader.IsDBNull(6) ? null : reader.GetString(6),
                        Severite = Enum.TryParse<SeveriteLog>(reader.GetString(7), out var sev) ? sev : SeveriteLog.Info
                    });
                }
            }

            var result = new List<SessionJournal>(sessions.Values);
            result.Sort((a, b) => b.Debut.CompareTo(a.Debut));
            return result;
        }

        public async Task<List<string>> ChargerListeUtilisateursAsync()
        {
            var users = new List<string>();
            using var c = new SqliteConnection(_connectionString);
            await c.OpenAsync();
            using var cmd = c.CreateCommand();
            cmd.CommandText = "SELECT DISTINCT Utilisateur FROM Sessions ORDER BY Utilisateur;";
            using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync()) users.Add(reader.GetString(0));
            return users;
        }
    }
}
