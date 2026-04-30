using System;

namespace Metrologo.Models
{
    public enum RoleUtilisateur
    {
        Utilisateur,
        Administrateur
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
        /// Affichage convivial (Prénom Nom) pour l'UI. Tombe en arrière sur le login
        /// si Nom/Prénom ne sont pas renseignés (cas des comptes hérités sans ces champs).
        /// </summary>
        public string NomComplet =>
            !string.IsNullOrWhiteSpace(Prenom) || !string.IsNullOrWhiteSpace(Nom)
                ? $"{Prenom} {Nom}".Trim()
                : Login;
    }
}
