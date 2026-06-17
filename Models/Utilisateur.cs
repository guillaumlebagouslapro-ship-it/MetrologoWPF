using System;

namespace Metrologo.Models
{
    public enum RoleUtilisateur
    {
        Utilisateur,
        Administrateur,
        /// <summary>
        /// Le "chef" : il peut tout ce que fait un Administrateur, et en plus il gère les rôles et le
        /// mot de passe admin. Il en reste toujours au moins un — impossible de supprimer ni de
        /// rétrograder le dernier.
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
        /// Le hash du mot de passe qui donne accès à la zone admin. Vaut null pour les simples Utilisateurs.
        /// On le renseigne à la promotion en Admin/SuperAdmin, et on l'efface à la rétrogradation.
        /// </summary>
        public string? PasswordHash { get; set; }

        /// <summary>Affichage "Prénom Nom" pour l'UI ; à défaut, on retombe sur le login pour les vieux comptes qui n'ont pas ces champs.</summary>
        public string NomComplet =>
            !string.IsNullOrWhiteSpace(Prenom) || !string.IsNullOrWhiteSpace(Nom)
                ? $"{Prenom} {Nom}".Trim()
                : Login;
    }
}
