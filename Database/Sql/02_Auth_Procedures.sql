-- =============================================================================
-- Procédures stockées d'authentification
-- =============================================================================
-- Centralisées côté SQL Server pour pouvoir évoluer (lock après N tentatives,
-- audit, etc.) sans toucher au code C#. Le hash PBKDF2 est calculé côté
-- application — la base ne voit jamais le mot de passe en clair.
-- =============================================================================

IF OBJECT_ID('dbo.SP_Utilisateur_Authentifier', 'P') IS NOT NULL
    DROP PROCEDURE dbo.SP_Utilisateur_Authentifier;
GO

CREATE PROCEDURE dbo.SP_Utilisateur_Authentifier
    @Login         NVARCHAR(50),
    @PasswordHash  NVARCHAR(200)
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @Id    INT;
    DECLARE @Role  NVARCHAR(20);
    DECLARE @Nom   NVARCHAR(100);
    DECLARE @Prenom NVARCHAR(100);

    SELECT TOP 1
        @Id = Id,
        @Role = Role,
        @Nom = Nom,
        @Prenom = Prenom
    FROM dbo.T_UTILISATEURS
    WHERE Login = @Login
      AND PasswordHash = @PasswordHash
      AND Actif = 1;

    IF @Id IS NOT NULL
    BEGIN
        -- Trace du login pour audit (non bloquant : si la mise à jour échoue,
        -- l'auth est quand même validée — c'est juste de la métadonnée).
        UPDATE dbo.T_UTILISATEURS
        SET DernierLogin = SYSUTCDATETIME()
        WHERE Id = @Id;

        SELECT @Id AS Id, @Login AS Login, @Nom AS Nom, @Prenom AS Prenom, @Role AS Role;
    END
END
GO

-- -----------------------------------------------------------------------------
-- Récupération hash courant (utile pour vérifier le mot de passe sans renvoyer
-- l'utilisateur — pour le futur changement de mot de passe par l'utilisateur).
-- -----------------------------------------------------------------------------

IF OBJECT_ID('dbo.SP_Utilisateur_LirePasswordHash', 'P') IS NOT NULL
    DROP PROCEDURE dbo.SP_Utilisateur_LirePasswordHash;
GO

CREATE PROCEDURE dbo.SP_Utilisateur_LirePasswordHash
    @Login NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;
    SELECT TOP 1 PasswordHash
    FROM dbo.T_UTILISATEURS
    WHERE Login = @Login AND Actif = 1;
END
GO
