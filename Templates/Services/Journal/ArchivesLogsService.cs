using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Dapper;
using Microsoft.Data.SqlClient;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Archivage automatique des logs : tous les mois, le mois précédent est exporté
    /// en fichiers JSON (un par jour) puis supprimé de la base SQL Server. Allège SQL
    /// pour qu'il garde des perfs constantes même après des années d'utilisation.
    ///
    /// Structure d'archives :
    /// <code>
    /// %ProgramData%\Metrologo\Archives\
    ///   2026-04\
    ///     2026-04-01\logs.json
    ///     2026-04-02\logs.json
    ///     ...
    ///   2026-03\
    ///     2026-03-01\logs.json
    ///     ...
    /// </code>
    ///
    /// Déclenchement : au démarrage de l'app — si le mois précédent n'est pas encore
    /// archivé (dossier absent), on archive. Idempotent : appel répété sans effet si
    /// déjà fait. Concurrence multi-postes gérée via verrou SQL <c>sp_getapplock</c> :
    /// si Baie et Paillasse démarrent en même temps le 1er du mois, un seul des deux
    /// archive, l'autre voit "déjà fait" et passe.
    ///
    /// Stockage par défaut : <c>%ProgramData%</c> (partagé entre utilisateurs Windows).
    /// En multi-postes, surcharge possible via la variable d'environnement
    /// <c>METROLOGO_ARCHIVES_PATH</c> pour pointer vers un partage réseau.
    /// </summary>
    public static class ArchivesLogsService
    {
        private const string LOCK_RESOURCE = "metrologo_archive_logs";
        private const string SOUS_DOSSIER_DEFAUT = "Metrologo\\Archives";
        private const string NOM_VAR_ENV = "METROLOGO_ARCHIVES_PATH";
        private const string NOM_FICHIER_CONFIG = "archives.path.txt";

        /// <summary>
        /// Chemin du dossier d'archives, résolu dans cet ordre :
        ///   1. Fichier <c>%LocalAppData%\Metrologo\archives.path.txt</c> (config par poste,
        ///      modifiable à chaud — c'est ce qu'on utilise pour pointer tous les postes
        ///      vers un partage réseau commun, ex. <c>\\PC-SERVEUR\MetrologoArchives</c>).
        ///   2. Variable d'environnement <c>METROLOGO_ARCHIVES_PATH</c> (override CI/CD).
        ///   3. Défaut local <c>%ProgramData%\Metrologo\Archives</c> (dev mono-poste).
        ///
        /// En prod multi-postes : configurer (1) avec un partage réseau accessible à tous
        /// les postes pour que l'admin puisse consulter les archives depuis n'importe quelle
        /// machine — l'archivage écrit, la consultation lit, même chemin partout.
        /// </summary>
        public static string DossierArchivesRacine
        {
            get
            {
                // 1. Fichier de config par poste (priorité max — modifiable à chaud)
                try
                {
                    string cheminConf = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "Metrologo", NOM_FICHIER_CONFIG);
                    if (File.Exists(cheminConf))
                    {
                        string contenu = File.ReadAllText(cheminConf).Trim();
                        if (!string.IsNullOrWhiteSpace(contenu)) return contenu;
                    }
                }
                catch { /* fallback silencieux */ }

                // 2. Variable d'environnement
                string custom = Environment.GetEnvironmentVariable(NOM_VAR_ENV) ?? "";
                if (!string.IsNullOrWhiteSpace(custom)) return custom;

                // 3. Défaut local
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    SOUS_DOSSIER_DEFAUT);
            }
        }

        /// <summary>
        /// Définit le chemin du dossier d'archives en l'écrivant dans le fichier de
        /// config par poste. À utiliser depuis l'UI admin pour pointer vers un partage
        /// réseau (ex. <c>\\PC-SERVEUR\MetrologoArchives</c>) sans manipuler de
        /// variables d'environnement.
        /// </summary>
        public static void DefinirDossierArchives(string chemin)
        {
            string dossier = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Metrologo");
            Directory.CreateDirectory(dossier);
            File.WriteAllText(Path.Combine(dossier, NOM_FICHIER_CONFIG), chemin?.Trim() ?? "");
        }

        // -------------------------------------------------------------------------
        // Déclencheurs
        // -------------------------------------------------------------------------

        /// <summary>
        /// Au démarrage de l'app : archive le mois précédent s'il n'est pas déjà archivé.
        /// No-op si la base / le serveur n'est pas joignable (l'app continue normalement).
        /// </summary>
        public static async Task ArchiverMoisPrecedentSiNecessaireAsync()
        {
            try
            {
                var maintenantUtc = DateTime.UtcNow;
                // Premier instant du mois précédent (UTC).
                var debutMoisPrec = new DateTime(maintenantUtc.Year, maintenantUtc.Month, 1, 0, 0, 0, DateTimeKind.Utc).AddMonths(-1);
                await ArchiverMoisInterneAsync(debutMoisPrec, force: false);
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Systeme, "ARCHIVE_AUTO_KO",
                    $"Archivage automatique au démarrage échoué : {ex.Message}");
            }
        }

        /// <summary>
        /// Archive un mois donné, force le ré-archivage même si le dossier existe.
        /// Utilisé par le bouton admin "Archiver maintenant".
        /// </summary>
        public static async Task<int> ArchiverMoisAsync(DateTime mois, bool force = true)
        {
            var debutMois = new DateTime(mois.Year, mois.Month, 1, 0, 0, 0, DateTimeKind.Utc);
            return await ArchiverMoisInterneAsync(debutMois, force);
        }

        // -------------------------------------------------------------------------
        // Liste / lecture des archives existantes
        // -------------------------------------------------------------------------

        public static List<DateTime> ListerMoisArchives()
        {
            var racine = DossierArchivesRacine;
            if (!Directory.Exists(racine)) return new List<DateTime>();
            var mois = new List<DateTime>();
            foreach (var dir in Directory.GetDirectories(racine))
            {
                string nom = Path.GetFileName(dir);
                if (DateTime.TryParseExact(nom, "yyyy-MM", CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out var dt))
                    mois.Add(dt);
            }
            return mois.OrderByDescending(d => d).ToList();
        }

        public static List<DateTime> ListerJoursDuMoisArchive(DateTime mois)
        {
            var dossier = Path.Combine(DossierArchivesRacine, mois.ToString("yyyy-MM"));
            if (!Directory.Exists(dossier)) return new List<DateTime>();
            var jours = new List<DateTime>();
            foreach (var dir in Directory.GetDirectories(dossier))
            {
                string nom = Path.GetFileName(dir);
                if (DateTime.TryParseExact(nom, "yyyy-MM-dd", CultureInfo.InvariantCulture,
                        DateTimeStyles.None, out var dt))
                    jours.Add(dt);
            }
            return jours.OrderByDescending(d => d).ToList();
        }

        /// <summary>
        /// Charge les sessions d'un jour archivé depuis son fichier JSON. Retourne une
        /// liste compatible avec celle remontée par <see cref="SqlServerJournalService"/>
        /// pour que la UI ne fasse pas de différence entre logs en base et archivés.
        /// </summary>
        public static async Task<List<SessionJournal>> ChargerJournauxArchivesAsync(DateTime jour)
        {
            string fichier = Path.Combine(
                DossierArchivesRacine,
                jour.ToString("yyyy-MM"),
                jour.ToString("yyyy-MM-dd"),
                "logs.json");

            if (!File.Exists(fichier)) return new List<SessionJournal>();

            string json = await File.ReadAllTextAsync(fichier);
            var contenu = JsonSerializer.Deserialize<ArchiveContenu>(json);
            if (contenu == null) return new List<SessionJournal>();

            // Reconstruction SessionJournal avec timestamps remis en heure locale.
            var sessions = contenu.Sessions
                .Select(s => new SessionJournal
                {
                    SessionId = s.SessionId,
                    Utilisateur = s.Utilisateur,
                    Machine = s.Machine,
                    Poste = s.Poste,
                    Debut = DateTime.SpecifyKind(s.DebutUtc, DateTimeKind.Utc).ToLocalTime(),
                    Fin = s.FinUtc.HasValue
                        ? DateTime.SpecifyKind(s.FinUtc.Value, DateTimeKind.Utc).ToLocalTime()
                        : (DateTime?)null
                })
                .ToDictionary(s => s.SessionId);

            foreach (var e in contenu.Entries)
            {
                if (!sessions.TryGetValue(e.SessionId, out var s)) continue;
                s.Entrees.Add(new LogEntry
                {
                    EntryId = e.EntryId,
                    SessionId = e.SessionId,
                    Timestamp = DateTime.SpecifyKind(e.TimestampUtc, DateTimeKind.Utc).ToLocalTime(),
                    Categorie = Enum.TryParse<CategorieLog>(e.Categorie, out var cat) ? cat : CategorieLog.Systeme,
                    Action = e.Action,
                    Message = e.Message ?? "",
                    Details = e.Details,
                    Severite = Enum.TryParse<SeveriteLog>(e.Severite, out var sev) ? sev : SeveriteLog.Info
                });
            }

            return sessions.Values.OrderByDescending(s => s.Debut).ToList();
        }

        // -------------------------------------------------------------------------
        // Implémentation
        // -------------------------------------------------------------------------

        private static async Task<int> ArchiverMoisInterneAsync(DateTime debutMoisUtc, bool force)
        {
            string nomMois = debutMoisUtc.ToString("yyyy-MM");
            string dossierMois = Path.Combine(DossierArchivesRacine, nomMois);

            if (!force && Directory.Exists(dossierMois)) return 0;

            using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
            await c.OpenAsync();

            // Verrou applicatif : empêche que 2 postes archivent le même mois en parallèle.
            // Lock automatiquement libéré à la fermeture de la connexion (scope = Session).
            int lockResult = await c.ExecuteScalarAsync<int>(
                @"DECLARE @r INT;
                  EXEC @r = sp_getapplock @Resource = @Res, @LockMode = 'Exclusive', @LockTimeout = 5000;
                  SELECT @r",
                new { Res = LOCK_RESOURCE });

            if (lockResult < 0)
            {
                JournalLog.Warn(CategorieLog.Systeme, "ARCHIVE_LOCK_TIMEOUT",
                    "Verrou archivage non obtenu (autre poste en cours d'archivage).");
                return 0;
            }

            // Re-vérification après lock (un autre poste a pu finir entre temps).
            if (!force && Directory.Exists(dossierMois)) return 0;

            var debutUtc = debutMoisUtc;
            var finUtc = debutMoisUtc.AddMonths(1);

            // 1. Lecture des entries du mois
            var entries = (await c.QueryAsync<EntryRowSql>(
                @"SELECT EntryId, SessionId, Timestamp, Categorie, Action, Message, Details, Severite
                  FROM dbo.T_LOGS_ENTRIES
                  WHERE Timestamp >= @Debut AND Timestamp < @Fin",
                new { Debut = debutUtc, Fin = finUtc })).ToList();

            // 2. Lecture des sessions du mois
            var sessions = (await c.QueryAsync<SessionRowSql>(
                @"SELECT SessionId, Utilisateur, Machine, Poste, Debut, Fin
                  FROM dbo.T_LOGS_SESSIONS
                  WHERE Debut >= @Debut AND Debut < @Fin",
                new { Debut = debutUtc, Fin = finUtc })).ToList();

            // Mois vide = on crée le dossier vide (marqueur d'archivage tenté).
            if (entries.Count == 0 && sessions.Count == 0)
            {
                Directory.CreateDirectory(dossierMois);
                JournalLog.Info(CategorieLog.Systeme, "ARCHIVE_MOIS_VIDE",
                    $"Aucun log à archiver pour {nomMois}. Dossier vide créé comme marqueur.");
                return 0;
            }

            // 3. Groupement par jour LOCAL (cohérence avec l'affichage UI)
            var sessionsParId = sessions.ToDictionary(s => s.SessionId);
            var entriesParJourLocal = entries
                .GroupBy(e => DateTime.SpecifyKind(e.Timestamp, DateTimeKind.Utc).ToLocalTime().Date)
                .ToDictionary(g => g.Key, g => g.ToList());
            var sessionsParJourLocal = sessions
                .GroupBy(s => DateTime.SpecifyKind(s.Debut, DateTimeKind.Utc).ToLocalTime().Date)
                .ToDictionary(g => g.Key, g => g.ToList());

            var tousJours = new HashSet<DateTime>();
            foreach (var k in entriesParJourLocal.Keys) tousJours.Add(k);
            foreach (var k in sessionsParJourLocal.Keys) tousJours.Add(k);

            Directory.CreateDirectory(dossierMois);

            int totalEntries = 0;
            foreach (var jour in tousJours.OrderBy(j => j))
            {
                string nomJour = jour.ToString("yyyy-MM-dd");
                string dossierJour = Path.Combine(dossierMois, nomJour);
                Directory.CreateDirectory(dossierJour);

                var sessionsDuJour = sessionsParJourLocal.TryGetValue(jour, out var sj)
                    ? sj : new List<SessionRowSql>();
                var entriesDuJour = entriesParJourLocal.TryGetValue(jour, out var ej)
                    ? ej : new List<EntryRowSql>();

                // Inclure les sessions parents des entries qui ne tombent pas dans ce jour
                // (session ouverte la veille, encore active aujourd'hui — son entry du jour
                // référence une session dont le Debut est antérieur).
                var idsManquants = entriesDuJour.Select(e => e.SessionId).Distinct()
                    .Except(sessionsDuJour.Select(s => s.SessionId)).ToList();
                foreach (var sid in idsManquants)
                    if (sessionsParId.TryGetValue(sid, out var s))
                        sessionsDuJour.Add(s);

                var contenu = new ArchiveContenu
                {
                    Date = jour,
                    Sessions = sessionsDuJour.Select(s => new SessionArchive
                    {
                        SessionId = s.SessionId,
                        Utilisateur = s.Utilisateur,
                        Machine = s.Machine,
                        Poste = s.Poste,
                        DebutUtc = DateTime.SpecifyKind(s.Debut, DateTimeKind.Utc),
                        FinUtc = s.Fin.HasValue ? DateTime.SpecifyKind(s.Fin.Value, DateTimeKind.Utc) : (DateTime?)null
                    }).ToList(),
                    Entries = entriesDuJour
                        .OrderBy(e => e.Timestamp)
                        .Select(e => new EntryArchive
                        {
                            EntryId = e.EntryId,
                            SessionId = e.SessionId,
                            TimestampUtc = DateTime.SpecifyKind(e.Timestamp, DateTimeKind.Utc),
                            Categorie = e.Categorie,
                            Action = e.Action,
                            Message = e.Message,
                            Details = e.Details,
                            Severite = e.Severite
                        }).ToList()
                };

                string fichierJour = Path.Combine(dossierJour, "logs.json");
                await File.WriteAllTextAsync(fichierJour,
                    JsonSerializer.Serialize(contenu, new JsonSerializerOptions { WriteIndented = true }));

                totalEntries += contenu.Entries.Count;
            }

            // 4. Suppression des logs archivés de la base (entries d'abord — FK).
            await c.ExecuteAsync(
                "DELETE FROM dbo.T_LOGS_ENTRIES WHERE Timestamp >= @Debut AND Timestamp < @Fin",
                new { Debut = debutUtc, Fin = finUtc });

            // Sessions : ne supprimer que celles qui n'ont plus aucune entry rattachée (au cas
            // où une session ouverte le 31 mars contiendrait encore des entries le 1er avril).
            await c.ExecuteAsync(
                @"DELETE FROM dbo.T_LOGS_SESSIONS
                  WHERE Debut >= @Debut AND Debut < @Fin
                    AND NOT EXISTS (SELECT 1 FROM dbo.T_LOGS_ENTRIES e WHERE e.SessionId = T_LOGS_SESSIONS.SessionId)",
                new { Debut = debutUtc, Fin = finUtc });

            JournalLog.Info(CategorieLog.Systeme, "ARCHIVE_LOGS_OK",
                $"Logs archivés : {nomMois} ({totalEntries} entrée(s), {sessions.Count} session(s)). " +
                $"Dossier : {dossierMois}");

            return totalEntries;
        }

        // -------------------------------------------------------------------------
        // DTOs (Dapper SQL ↔ JSON disque)
        // -------------------------------------------------------------------------

        private sealed class EntryRowSql
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

        private sealed class SessionRowSql
        {
            public string SessionId { get; set; } = "";
            public string Utilisateur { get; set; } = "";
            public string Machine { get; set; } = "";
            public string? Poste { get; set; }
            public DateTime Debut { get; set; }
            public DateTime? Fin { get; set; }
        }

        public class ArchiveContenu
        {
            public DateTime Date { get; set; }
            public List<SessionArchive> Sessions { get; set; } = new();
            public List<EntryArchive> Entries { get; set; } = new();
        }

        public class SessionArchive
        {
            public string SessionId { get; set; } = "";
            public string Utilisateur { get; set; } = "";
            public string Machine { get; set; } = "";
            public string? Poste { get; set; }
            public DateTime DebutUtc { get; set; }
            public DateTime? FinUtc { get; set; }
        }

        public class EntryArchive
        {
            public long EntryId { get; set; }
            public string SessionId { get; set; } = "";
            public DateTime TimestampUtc { get; set; }
            public string Categorie { get; set; } = "";
            public string Action { get; set; } = "";
            public string? Message { get; set; }
            public string? Details { get; set; }
            public string Severite { get; set; } = "";
        }
    }
}
