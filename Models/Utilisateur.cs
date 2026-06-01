using System;

namespace Metrologo.Models
{
    public enum RoleUtilisateur
    {
        Utilisateur,
        Administrateur,
        /// <summary>
        /// « Chef » — peut tout ce que fait un Administrateur, ET en plus :
        /// modifier les rôles des autres utilisateurs et changer le mot de passe
        /// d'accès à la zone admin. Il y a toujours au moins un SuperAdministrateur
        /// (impossible de supprimer ou de rétrograder le dernier).
        /// </summary>
        SuperAdministrateur
    }

    public class Utilisateur
    {
        public int Id { get; set; }
        public string Login { get; set; } = string.Empty;
        public string Nom { get; set; } = string.Empty;
        public string Prenom { get; set; } = string.Empty;
        public RoleUtilisateur Role { get; set; }
        public bool Actif { get; set; } = true;
        public DateTime DateCreation { get; set; }
        public DateTime? DernierLogin { get; set; }

        /// <summary>
        /// Hash du mot de passe d'accès à la zone admin. <c>null</c> pour les comptes
        /// <see cref="RoleUtilisateur.Utilisateur"/> (ils ne se connectent jamais à l'admin).
        /// Renseigné automatiquement à la promotion en Admin/SuperAdmin et effacé à
        /// la rétrogradation.
        /// </summary>
        public string? PasswordHash { get; set; }

        /// <summary>
        /// Affichage convivial (Prénom Nom) pour l'UI. Tombe en arrière sur le login
        /// si Nom/Prénom ne sont pas renseignés (cas des comptes hérités sans ces champs).
        /// </summary>
        public string NomComplet =>
            !string.IsNullOrWhiteSpace(Prenom) || !string.IsNullOrWhiteSpace(Nom)
                ? $"{Prenom} {Nom}".Trim()
                : Login;
    }
}
