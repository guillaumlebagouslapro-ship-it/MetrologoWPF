using System;

namespace Metrologo.Models
{
    public enum RoleUtilisateur
    {
        Utilisateur,
        Administrateur,
        /// <summary>
        /// Le "chef" : tout ce que fait un Administrateur, plus la gestion des rôles et du mot de
        /// passe admin. Il en reste toujours au moins un (impossible de supprimer ou rétrograder le dernier).
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
        /// Hash du mot de passe d'accès à la zone admin. null pour les simples Utilisateurs.
        /// Renseigné à la promotion en Admin/SuperAdmin, effacé à la rétrogradation.
        /// </summary>
        public string? PasswordHash { get; set; }

        /// <summary>Affichage "Prénom Nom" pour l'UI ; repli sur le login pour les comptes hérités sans ces champs.</summary>
        public string NomComplet =>
            !string.IsNullOrWhiteSpace(Prenom) || !string.IsNullOrWhiteSpace(Nom)
                ? $"{Prenom} {Nom}".Trim()
                : Login;
    }
}
