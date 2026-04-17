namespace Metrologo.Models
{
    public class Rubidium
    {
        public int Id { get; set; }
        public string Designation { get; set; } = string.Empty;
        public double FrequenceMoyenne { get; set; }
        public bool AvecGPS { get; set; }

        // Une petite propriété pratique qui formate le nom tout seul
        public string NomAffichage => AvecGPS ? $"GPS + {Designation}" : Designation;
    }
}