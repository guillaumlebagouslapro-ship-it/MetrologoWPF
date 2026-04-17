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
        public RoleUtilisateur Role { get; set; }
    }
}