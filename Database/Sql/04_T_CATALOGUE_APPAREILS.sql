-- =============================================================================
-- Catalogue centralisé des modèles d'appareils SCPI
-- =============================================================================
-- Remplace l'ancien JSON local (un fichier par poste, divergence garantie) par
-- une table SQL Server centralisée que tous les postes consultent. Création /
-- modification d'un modèle sur un poste = visible immédiatement sur les autres.
--
-- Stratégie de stockage : colonnes plates pour les champs cherchables (Nom,
-- IDN, dates), colonne JSON `Configuration` pour le reste (Parametres SCPI,
-- Gates, Entrees, Couplages, Reglages). Évite de multiplier les tables liées
-- pour des collections qu'on ne joint jamais — la complexité des Reglages
-- (avec sous-options et commandes SCPI imbriquées) ne se prête pas au
-- relationnel pur, et le schéma reste évolutif sans migration.
-- =============================================================================

IF OBJECT_ID('dbo.T_CATALOGUE_APPAREILS', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.T_CATALOGUE_APPAREILS
    (
        Id            NVARCHAR(80)   NOT NULL,        -- slug stable (ex: "agilent-53131a-7f3a")
        Nom           NVARCHAR(150)  NOT NULL,        -- nom lisible
        FabricantIdn  NVARCHAR(100)  NOT NULL,        -- pour matching IDN au scan
        ModeleIdn     NVARCHAR(100)  NOT NULL,
        Configuration NVARCHAR(MAX)  NOT NULL,        -- JSON : Parametres + Gates + Entrees + Couplages + Reglages
        DateCreation  DATETIME2(0)   NOT NULL CONSTRAINT DF_T_CAT_APP_DateCreation DEFAULT (SYSUTCDATETIME()),
        CreePar       NVARCHAR(50)   NOT NULL CONSTRAINT DF_T_CAT_APP_CreePar DEFAULT (''),
        DateModif     DATETIME2(0)   NOT NULL CONSTRAINT DF_T_CAT_APP_DateModif DEFAULT (SYSUTCDATETIME()),
        CONSTRAINT PK_T_CATALOGUE_APPAREILS PRIMARY KEY CLUSTERED (Id)
    );

    CREATE INDEX IX_T_CATALOGUE_APPAREILS_Idn ON dbo.T_CATALOGUE_APPAREILS (FabricantIdn, ModeleIdn);
    CREATE INDEX IX_T_CATALOGUE_APPAREILS_Nom ON dbo.T_CATALOGUE_APPAREILS (Nom);
END
GO
