-- =============================================================================
-- T_UTILISATEURS — table d'authentification Metrologo
-- =============================================================================
-- Schéma centralisé sur SQL Server : un seul endroit pour les comptes utilisateur,
-- accédé par tous les postes (Baie / Paillasse / dev). L'authentification se fait
-- côté application via PBKDF2 ; la base ne stocke que le hash + sel encodés.
--
-- Convention de hash (cf. PasswordHasher.cs) :
--   PasswordHash = "PBKDF2$<iterations>$<salt-base64>$<hash-base64>"
--   ex : "PBKDF2$100000$d3hLN2Z...==$Ab12Yz...=="
-- Le format auto-décrit permet d'augmenter le nombre d'itérations dans le futur
-- sans casser les anciens comptes (le hasher choisit en relisant la chaîne).
--
-- Le login est dérivé automatiquement de prenom.nom au moment de la création
-- (côté application) — l'admin ne saisit que Nom + Prénom.
-- =============================================================================

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
END
GO

-- -----------------------------------------------------------------------------
-- Compte administrateur par défaut
-- -----------------------------------------------------------------------------
-- Login : admin
-- Mdp   : admin123  (à changer au premier login en prod)
--
-- Hash généré via PasswordHasher.HashPassword("admin123") :
--   PBKDF2$100000$<sel>$<hash>
-- Le hash ci-dessous est un placeholder — il sera remplacé au moment de l'exec
-- par DatabaseInitializer si la ligne admin n'existe pas (méthode plus sûre que
-- de figer un hash en clair dans le SQL versionné).
-- -----------------------------------------------------------------------------

-- L'insertion de l'admin par défaut est gérée côté C# (DatabaseInitializer)
-- pour éviter de versionner un hash et permettre le re-hash si l'algo évolue.
