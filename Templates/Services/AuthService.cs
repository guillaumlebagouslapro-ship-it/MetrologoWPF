using System;
using System.Data;
using System.Threading.Tasks;
using Dapper;
using Metrologo.Models;
using Metrologo.Services.Journal;
using Microsoft.Data.SqlClient;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services
{
    public interface IAuthService
    {
        Task<Utilisateur?> AuthentifierAsync(string login, string password);
    }

    /// <summary>
    /// Authentification contre la base SQL Server centralisée Metrologo.
    /// Le hash PBKDF2 est calculé côté application (cf. <see cref="PasswordHasher"/>) :
    /// la base ne voit jamais le mot de passe en clair.
    ///
    /// Stratégie : on lit le hash stocké pour le login fourni, puis on vérifie en local
    /// avec <see cref="PasswordHasher.VerifyPassword"/> (à temps constant). Cette
    /// approche, contrairement à un <c>WHERE Login = ? AND PasswordHash = ?</c>
    /// direct, est nécessaire car PBKDF2 utilise un sel aléatoire par compte — donc
    /// le hash à comparer n'est connu qu'après lecture de la ligne.
    /// </summary>
    public class AuthService : IAuthService
    {
        public async Task<Utilisateur?> AuthentifierAsync(string login, string password)
        {
            if (string.IsNullOrWhiteSpace(login) || string.IsNullOrEmpty(password)) return null;

            try
            {
                using var connection = new SqlConnection(MetrologoDbConnection.ConnectionString);
                await connection.OpenAsync();

                // Étape 1 : on récupère hash stocké + métadonnées en une requête.
                // Si le compte n'existe pas / est inactif, on retourne null (l'appelant
                // ne distingue pas « login inconnu » de « mauvais mdp » — empêche
                // l'énumération des comptes valides).
                var utilisateur = await connection.QueryFirstOrDefaultAsync<UtilisateurAvecHash>(
                    @"SELECT TOP 1 Id, Login, Nom, Prenom, Role, PasswordHash
                      FROM dbo.T_UTILISATEURS
                      WHERE Login = @Login AND Actif = 1",
                    new { Login = login });

                if (utilisateur == null) return null;

                // Étape 2 : vérification PBKDF2 à temps constant.
                if (!PasswordHasher.VerifyPassword(password, utilisateur.PasswordHash))
                    return null;

                // Étape 3 : trace du login (best-effort, non bloquant).
                try
                {
                    await connection.ExecuteAsync(
                        "UPDATE dbo.T_UTILISATEURS SET DernierLogin = SYSUTCDATETIME() WHERE Id = @Id",
                        new { utilisateur.Id });
                }
                catch (Exception ex)
                {
                    JournalLog.Warn(CategorieLog.Authentification, "MAJ_DERNIER_LOGIN_ECHEC",
                        $"Mise à jour DernierLogin pour {utilisateur.Login} échouée : {ex.Message}");
                }

                return new Utilisateur
                {
                    Id = utilisateur.Id,
                    Login = utilisateur.Login,
                    Nom = utilisateur.Nom ?? string.Empty,
                    Prenom = utilisateur.Prenom ?? string.Empty,
                    Role = Enum.TryParse<RoleUtilisateur>(utilisateur.Role, out var r) ? r : RoleUtilisateur.Utilisateur
                };
            }
            catch (SqlException ex)
            {
                // SQL Server non joignable / instance absente / base manquante — on log
                // et on retourne null. Le LoginViewModel pourra retomber sur son fallback
                // local de phase BDD-déconnectée le temps que le serveur soit branché.
                JournalLog.Warn(CategorieLog.Authentification, "AUTH_SQL_ERR",
                    $"SQL Server inaccessible pour authentification : {ex.Message}");
                return null;
            }
        }

        // DTO interne — éviter d'exposer le PasswordHash dans le modèle public Utilisateur.
        private sealed class UtilisateurAvecHash
        {
            public int Id { get; set; }
            public string Login { get; set; } = string.Empty;
            public string? Nom { get; set; }
            public string? Prenom { get; set; }
            public string Role { get; set; } = "Utilisateur";
            public string PasswordHash { get; set; } = string.Empty;
        }
    }
}
