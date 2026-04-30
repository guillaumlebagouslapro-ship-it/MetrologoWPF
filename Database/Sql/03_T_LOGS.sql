-- =============================================================================
-- Journal d'activité Metrologo — tables centralisées sur SQL Server
-- =============================================================================
-- Remplace l'ancien stockage SQLite local (un fichier par poste, isolé) par une
-- base centralisée que l'admin peut consulter depuis n'importe quel poste pour
-- tracer l'activité de toutes les Baies / Paillasses.
--
-- Convention horaire : tout est stocké en UTC (DATETIME2 alimenté par
-- SYSUTCDATETIME() ou par les timestamps UTC envoyés depuis l'app). Conversion
-- en heure locale faite côté C# au moment de l'affichage (gère DST automatiquement).
-- =============================================================================

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
GO

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
END
GO
