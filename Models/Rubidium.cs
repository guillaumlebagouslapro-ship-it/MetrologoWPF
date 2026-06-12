namespace Metrologo.Models
{
    public class Rubidium
    {
        public int Id { get; set; }
        public string Designation { get; set; } = string.Empty;
        public double FrequenceMoyenne { get; set; }
        public bool AvecGPS { get; set; }

        /// <summary>Vrai si la fréquence a été saisie à la main par l'admin (aucun rubidium du
        /// catalogue sélectionné) : FrequenceMoyenne vient alors directement de la saisie.</summary>
        public bool EstReglageManuel { get; set; }

        // Affiche directement le nom choisi dans le catalogue (réglage manuel à part).
        public string NomAffichage => EstReglageManuel
            ? $"Réglage manuel ({Services.SaisieHelper.FormaterFrequence(FrequenceMoyenne)} Hz)"
            : Designation;

        /// <summary>Fréquence formatée pour l'affichage : groupée par milliers, précision exacte
        /// ("10 000 000" / "10 000 000,1234"). Pour les listes/onglets rubidium.</summary>
        public string FrequenceAffichee => Services.SaisieHelper.FormaterFrequence(FrequenceMoyenne);

        /// <summary>
        /// Libellé écrit dans la feuille de mesure (zone ZNRubidium) : la Designation du catalogue
        /// telle quelle, sans suffixe automatique (le nom est libre, y mettre "GPS" etc. si besoin).
        /// Pour un réglage manuel : "Réglage manuel (xxx Hz)".
        /// </summary>
        public string DesignationAvecRaccord => EstReglageManuel
            ? $"Réglage manuel ({FrequenceMoyenne} Hz)"
            : Designation;
    }
}