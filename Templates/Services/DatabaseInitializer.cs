using System;
using System.Threading.Tasks;
using Dapper;
using Metrologo.Services.Journal;
using Microsoft.Data.Sqlite;
using Microsoft.Data.SqlClient;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services
{
    /// <summary>
    /// Initialisation des bases de données du poste :
    ///   • SQLite locale  → tables encore en phase BDD déconnectée (rubidiums de test, etc.).
    ///   • SQL Server     → schéma utilisateurs (centralisé).
    ///
    /// Les deux initialisations sont indépendantes et best-effort : si SQL Server n'est
    /// pas joignable au démarrage (ex. serveur éteint, instance Express pas encore
    /// installée), la SQLite locale est quand même initialisée et l'app démarre.
    /// L'authentification retombera alors sur le fallback hardcodé du LoginViewModel
    /// (cf. TODO dans ce fichier — à retirer une fois le serveur en prod).
    /// </summary>
    public static class DatabaseInitializer
    {
        public static async Task InitialiserAsync()
        {
            await InitialiserSqliteLocaleAsync();
            await InitialiserSqlServerAsync();
        }

        // -------------------------------------------------------------------------
        // SQLite locale : tables encore stubées (rubidiums, etc.) — à retirer au fur
        // et à mesure que les emplacements TODO BDD sont branchés sur SQL Server.
        // -------------------------------------------------------------------------
        private static async Task InitialiserSqliteLocaleAsync()
        {
            using var connection = new SqliteConnection(DatabaseService.ConnectionString);
            await connection.OpenAsync();

            string createRubidium = @"
                CREATE TABLE IF NOT EXISTS T_RUBIDIUM (
                    Id               INTEGER PRIMARY KEY AUTOINCREMENT,
                    Designation      TEXT NOT NULL,
                    FrequenceMoyenne REAL NOT NULL DEFAULT 10.0,
                    AvecGPS          INTEGER NOT NULL DEFAULT 0
                );";

            using var cmd = new SqliteCommand(createRubidium, connection);
            await cmd.ExecuteNonQueryAsync();
        }

        // -------------------------------------------------------------------------
        // SQL Server centralisé : schéma utilisateurs + admin par défaut.
        // -------------------------------------------------------------------------
        private static async Task InitialiserSqlServerAsync()
        {
            try
            {
                using var connection = new SqlConnection(MetrologoDbConnection.ConnectionString);
                await connection.OpenAsync();

                // Création des tables si absentes (idempotent — IF NOT EXISTS).
                // On garde le DDL en dur ici pour que l'init marche même app packagée.
                // Les fichiers .sql versionnés dans Database/Sql/ restent la source de
                // vérité pour évolutions manuelles côté serveur central.
                string ddlUtilisateurs = @"
IF OBJECT_ID('dbo.T_UTILISATEURS', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.T_UTILISATEURS
    (
        Id            INT            IDENTITY(1,1) NOT NULL,
        Login         NVARCHAR(50)   NOT NULL,
        Nom           NVARCHAR(100)  NOT NULL,
        Prenom        NVARCHAR(100)  NOT NULL,
        PasswordHash  NVARCHAR(200)  NOT NULL,
        Role          NVARCHAR(20)   NOT NULL CONSTRAINT DF_T_UTILISATEURS_Role DEFAULT ('Utilisateur'),
        Actif         BIT            NOT NULL CONSTRAINT DF_T_UTILISATEURS_Actif DEFAULT (1),
        DateCreation  DATETIME2(0)   NOT NULL CONSTRAINT DF_T_UTILISATEURS_DateCreation DEFAULT (SYSUTCDATETIME()),
        DernierLogin  DATETIME2(0)   NULL,
        CONSTRAINT PK_T_UTILISATEURS PRIMARY KEY CLUSTERED (Id),
        CONSTRAINT UQ_T_UTILISATEURS_Login UNIQUE (Login),
        CONSTRAINT CK_T_UTILISATEURS_Role CHECK (Role IN ('Utilisateur','Administrateur'))
    );
    CREATE INDEX IX_T_UTILISATEURS_Actif ON dbo.T_UTILISATEURS (Actif) INCLUDE (Login, Role);
END";
                await connection.ExecuteAsync(ddlUtilisateurs);

                // Tables du journal centralisé (sessions + entries)
                string ddlLogs = @"
IF OBJECT_ID('dbo.T_LOGS_SESSIONS', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.T_LOGS_SESSIONS
    (
        SessionId    NVARCHAR(32)   NOT NULL,
        Utilisateur  NVARCHAR(50)   NOT NULL,
        Machine      NVARCHAR(100)  NOT NULL,
        Poste        NVARCHAR(20)   NULL,
        Debut        DATETIME2(0)   NOT NULL,
        Fin          DATETIME2(0)   NULL,
        CONSTRAINT PK_T_LOGS_SESSIONS PRIMARY KEY CLUSTERED (SessionId)
    );
    CREATE INDEX IX_T_LOGS_SESSIONS_Utilisateur ON dbo.T_LOGS_SESSIONS (Utilisateur);
    CREATE INDEX IX_T_LOGS_SESSIONS_Debut ON dbo.T_LOGS_SESSIONS (Debut DESC);
END
IF OBJECT_ID('dbo.T_LOGS_ENTRIES', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.T_LOGS_ENTRIES
    (
        EntryId      BIGINT         IDENTITY(1,1) NOT NULL,
        SessionId    NVARCHAR(32)   NOT NULL,
        Timestamp    DATETIME2(0)   NOT NULL,
        Categorie    NVARCHAR(30)   NOT NULL,
        Action       NVARCHAR(100)  NOT NULL,
        Message      NVARCHAR(MAX)  NULL,
        Details      NVARCHAR(MAX)  NULL,
        Severite     NVARCHAR(20)   NOT NULL,
        CONSTRAINT PK_T_LOGS_ENTRIES PRIMARY KEY CLUSTERED (EntryId),
        CONSTRAINT FK_T_LOGS_ENTRIES_SESSIONS FOREIGN KEY (SessionId)
            REFERENCES dbo.T_LOGS_SESSIONS (SessionId)
    );
    CREATE INDEX IX_T_LOGS_ENTRIES_Session ON dbo.T_LOGS_ENTRIES (SessionId);
    CREATE INDEX IX_T_LOGS_ENTRIES_Timestamp ON dbo.T_LOGS_ENTRIES (Timestamp DESC);
    CREATE INDEX IX_T_LOGS_ENTRIES_Action ON dbo.T_LOGS_ENTRIES (Action);
END";
                await connection.ExecuteAsync(ddlLogs);

                // Catalogue centralisé des modèles d'appareils SCPI
                string ddlCatalogue = @"
IF OBJECT_ID('dbo.T_CATALOGUE_APPAREILS', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.T_CATALOGUE_APPAREILS
    (
        Id            NVARCHAR(80)   NOT NULL,
        Nom           NVARCHAR(150)  NOT NULL,
        FabricantIdn  NVARCHAR(100)  NOT NULL,
        ModeleIdn     NVARCHAR(100)  NOT NULL,
        Configuration NVARCHAR(MAX)  NOT NULL,
        DateCreation  DATETIME2(0)   NOT NULL CONSTRAINT DF_T_CAT_APP_DateCreation DEFAULT (SYSUTCDATETIME()),
        CreePar       NVARCHAR(50)   NOT NULL CONSTRAINT DF_T_CAT_APP_CreePar DEFAULT (''),
        DateModif     DATETIME2(0)   NOT NULL CONSTRAINT DF_T_CAT_APP_DateModif DEFAULT (SYSUTCDATETIME()),
        CONSTRAINT PK_T_CATALOGUE_APPAREILS PRIMARY KEY CLUSTERED (Id)
    );
    CREATE INDEX IX_T_CATALOGUE_APPAREILS_Idn ON dbo.T_CATALOGUE_APPAREILS (FabricantIdn, ModeleIdn);
    CREATE INDEX IX_T_CATALOGUE_APPAREILS_Nom ON dbo.T_CATALOGUE_APPAREILS (Nom);
END";
                await connection.ExecuteAsync(ddlCatalogue);

                // Seed admin par défaut. Le hash est généré à la volée pour ne JAMAIS
                // figer une valeur en clair / en commit ; il est unique à chaque
                // déploiement (sel aléatoire) — donc même mot de passe par défaut,
                // hash différent par poste.
                int existeAdmin = await connection.ExecuteScalarAsync<int>(
                    "SELECT COUNT(*) FROM dbo.T_UTILISATEURS WHERE Login = 'admin'");

                if (existeAdmin == 0)
                {
                    string hash = PasswordHasher.HashPassword("admin123");
                    await connection.ExecuteAsync(
                        @"INSERT INTO dbo.T_UTILISATEURS (Login, Nom, Prenom, PasswordHash, Role)
                          VALUES ('admin', 'Admin', 'Metrologo', @Hash, 'Administrateur')",
                        new { Hash = hash });

                    JournalLog.Info(CategorieLog.Authentification, "ADMIN_SEED",
                        "Compte administrateur par défaut créé (admin/admin123) — à changer en prod.");
                }
            }
            catch (SqlException ex)
            {
                // SQL Server pas joignable au démarrage — on continue avec le fallback
                // hardcodé du LoginViewModel le temps que le serveur soit installé.
                JournalLog.Warn(CategorieLog.Authentification, "DB_INIT_SQL_INDISPO",
                    $"SQL Server indisponible à l'init ({ex.Message}). " +
                    "Fallback login local actif jusqu'à connexion serveur.");
            }
        }
    }
}
