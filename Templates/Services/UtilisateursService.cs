using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Dapper;
using Metrologo.Models;
using Metrologo.Services.Journal;
using Microsoft.Data.SqlClient;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services
{
    public interface IUtilisateursService
    {
        Task<IReadOnlyList<Utilisateur>> ListerAsync();
        Task<(Utilisateur utilisateur, string motDePasseClair)> CreerAsync(string nom, string prenom, RoleUtilisateur role);
        Task<bool> SupprimerAsync(int id);

        /// <summary>
        /// Génère un nouveau mot de passe aléatoire pour le compte donné, met à jour
        /// le hash en base, et retourne le mot de passe clair (à afficher une seule
        /// fois et communiquer à l'utilisateur). L'ancien mot de passe est immédiatement
        /// invalidé. Lance <see cref="InvalidOperationException"/> si l'utilisateur
        /// n'existe pas.
        /// </summary>
        Task<string> ReinitialiserMotDePasseAsync(int id);
    }

    /// <summary>
    /// CRUD utilisateurs sur SQL Server. La création génère automatiquement
    /// le login (prenom.nom + suffixe en cas de collision) et un mot de passe
    /// aléatoire — l'admin ne saisit que Nom et Prénom. Le mot de passe clair
    /// n'est retourné qu'une seule fois (au moment de la création) pour être
    /// communiqué à l'utilisateur ; il n'est jamais stocké en clair.
    /// </summary>
    public class UtilisateursService : IUtilisateursService
    {
        public async Task<IReadOnlyList<Utilisateur>> ListerAsync()
        {
            using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
            await c.OpenAsync();

            var rows = await c.QueryAsync<UtilisateurRow>(
                @"SELECT Id, Login, Nom, Prenom, Role, Actif, DateCreation, DernierLogin
                  FROM dbo.T_UTILISATEURS
                  ORDER BY Nom, Prenom");

            var liste = new List<Utilisateur>();
            foreach (var r in rows)
            {
                liste.Add(new Utilisateur
                {
                    Id = r.Id,
                    Login = r.Login,
                    Nom = r.Nom ?? "",
                    Prenom = r.Prenom ?? "",
                    Role = Enum.TryParse<RoleUtilisateur>(r.Role, out var ro) ? ro : RoleUtilisateur.Utilisateur,
                    Actif = r.Actif,
                    // Les colonnes DATETIME2 sont alimentées par SYSUTCDATETIME() côté SQL,
                    // mais Dapper les remonte en DateTimeKind.Unspecified — sans conversion,
                    // l'affichage UI montrerait l'heure UTC (décalage de 2h en CEST). On marque
                    // Kind=Utc puis on convertit en heure locale du poste.
                    DateCreation = DateTime.SpecifyKind(r.DateCreation, DateTimeKind.Utc).ToLocalTime(),
                    DernierLogin = r.DernierLogin.HasValue
                        ? DateTime.SpecifyKind(r.DernierLogin.Value, DateTimeKind.Utc).ToLocalTime()
                        : (DateTime?)null
                });
            }
            return liste;
        }

        public async Task<(Utilisateur utilisateur, string motDePasseClair)> CreerAsync(
            string nom, string prenom, RoleUtilisateur role)
        {
            if (string.IsNullOrWhiteSpace(nom)) throw new ArgumentException("Nom requis", nameof(nom));
            if (string.IsNullOrWhiteSpace(prenom)) throw new ArgumentException("Prénom requis", nameof(prenom));

            using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
            await c.OpenAsync();

            // Login unique : on tente prenom.nom puis prenom.nom2, .nom3, etc.
            string baseLogin = LoginGenerator.GenererBase(prenom, nom);
            string login = await TrouverLoginDisponibleAsync(c, baseLogin);

            string motDePasseClair = MotDePasseGenerator.Generer();
            string hash = PasswordHasher.HashPassword(motDePasseClair);

            int id = await c.ExecuteScalarAsync<int>(
                @"INSERT INTO dbo.T_UTILISATEURS (Login, Nom, Prenom, PasswordHash, Role)
                  OUTPUT INSERTED.Id
                  VALUES (@Login, @Nom, @Prenom, @Hash, @Role)",
                new { Login = login, Nom = nom.Trim(), Prenom = prenom.Trim(), Hash = hash, Role = role.ToString() });

            JournalLog.Info(CategorieLog.Administration, "UTILISATEUR_CREE",
                $"Compte créé : {login} ({prenom} {nom}, {role}).");

            return (new Utilisateur
            {
                Id = id,
                Login = login,
                Nom = nom.Trim(),
                Prenom = prenom.Trim(),
                Role = role,
                Actif = true,
                DateCreation = DateTime.UtcNow
            }, motDePasseClair);
        }

        public async Task<bool> SupprimerAsync(int id)
        {
            using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
            await c.OpenAsync();

            // Garde-fou : ne jamais supprimer le dernier admin (sinon plus personne
            // ne peut administrer la base). On vérifie côté app — la contrainte
            // pourrait aussi être en TRIGGER côté SQL Server pour être à toute épreuve.
            var cible = await c.QueryFirstOrDefaultAsync<UtilisateurRow>(
                "SELECT Id, Login, Role FROM dbo.T_UTILISATEURS WHERE Id = @Id",
                new { Id = id });
            if (cible == null) return false;

            if (string.Equals(cible.Role, "Administrateur", StringComparison.OrdinalIgnoreCase))
            {
                int nbAdminsActifs = await c.ExecuteScalarAsync<int>(
                    "SELECT COUNT(*) FROM dbo.T_UTILISATEURS WHERE Role = 'Administrateur' AND Actif = 1");
                if (nbAdminsActifs <= 1)
                {
                    JournalLog.Warn(CategorieLog.Administration, "SUPPRESSION_REFUSEE",
                        $"Refus de supprimer {cible.Login} : dernier administrateur actif.");
                    throw new InvalidOperationException(
                        "Impossible de supprimer le dernier administrateur actif.");
                }
            }

            int affected = await c.ExecuteAsync(
                "DELETE FROM dbo.T_UTILISATEURS WHERE Id = @Id", new { Id = id });

            if (affected > 0)
            {
                JournalLog.Info(CategorieLog.Administration, "UTILISATEUR_SUPPRIME",
                    $"Compte supprimé : {cible.Login} (Id={id}).");
            }
            return affected > 0;
        }

        public async Task<string> ReinitialiserMotDePasseAsync(int id)
        {
            using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
            await c.OpenAsync();

            var cible = await c.QueryFirstOrDefaultAsync<UtilisateurRow>(
                "SELECT Id, Login FROM dbo.T_UTILISATEURS WHERE Id = @Id", new { Id = id });
            if (cible == null)
                throw new InvalidOperationException($"Utilisateur Id={id} introuvable.");

            string nouveauMdpClair = MotDePasseGenerator.Generer();
            string hash = PasswordHasher.HashPassword(nouveauMdpClair);

            await c.ExecuteAsync(
                "UPDATE dbo.T_UTILISATEURS SET PasswordHash = @Hash WHERE Id = @Id",
                new { Hash = hash, Id = id });

            JournalLog.Warn(CategorieLog.Administration, "MDP_REINITIALISE",
                $"Mot de passe réinitialisé pour {cible.Login} (Id={id}).");

            return nouveauMdpClair;
        }

        private static async Task<string> TrouverLoginDisponibleAsync(SqlConnection c, string baseLogin)
        {
            // 1ère tentative : login nu. Puis suffixe numérique 2, 3, 4… si pris.
            string candidat = baseLogin;
            for (int suffixe = 2; suffixe < 1000; suffixe++)
            {
                int existe = await c.ExecuteScalarAsync<int>(
                    "SELECT COUNT(*) FROM dbo.T_UTILISATEURS WHERE Login = @L",
                    new { L = candidat });
                if (existe == 0) return candidat;
                candidat = $"{baseLogin}{suffixe}";
            }
            // Dernier recours : suffixe aléatoire (improbable d'arriver ici).
            return $"{baseLogin}{Guid.NewGuid().ToString("N")[..6]}";
        }

        private sealed class UtilisateurRow
        {
            public int Id { get; set; }
            public string Login { get; set; } = "";
            public string? Nom { get; set; }
            public string? Prenom { get; set; }
            public string Role { get; set; } = "Utilisateur";
            public bool Actif { get; set; }
            public DateTime DateCreation { get; set; }
            public DateTime? DernierLogin { get; set; }
        }
    }
}
