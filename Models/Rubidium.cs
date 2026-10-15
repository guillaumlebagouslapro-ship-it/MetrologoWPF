namespace Metrologo.Models
{
    public class Rubidium
    {
        public int Id { get; set; }
        public string Designation { get; set; } = string.Empty;
        public double FrequenceMoyenne { get; set; }
        public bool AvecGPS { get; set; }

        /// <summary>Vrai quand l'admin a saisi la fréquence à la main (sans choisir de rubidium dans le
        /// catalogue) : dans ce cas, FrequenceMoyenne reprend directement ce qui a été saisi.</summary>
        public bool EstReglageManuel { get; set; }

        // On affiche tel quel le nom choisi dans le catalogue (le réglage manuel étant un cas à part).
        public string NomAffichage => EstReglageManuel
            ? $"Réglage manuel ({Services.SaisieHelper.FormaterFrequence(FrequenceMoyenne)} Hz)"
            : Designation;

        /// <summary>La fréquence mise en forme pour l'affichage : groupée par milliers et sans perte de
        /// précision ("10 000 000" / "10 000 000,1234"). Sert dans les listes et onglets rubidium.</summary>
        public string FrequenceAffichee => Services.SaisieHelper.FormaterFrequence(FrequenceMoyenne);

        /// <summary>
        /// Libellé écrit dans la feuille de mesure (zone ZNRubidium) : la Designation du catalogue
        /// telle quelle (rien n'est ajouté, "GPS" etc. se saisit dans le nom). Réglage manuel :
        /// "Réglage manuel (xxx Hz)".
        /// </summary>
        public string DesignationAvecRaccord => EstReglageManuel
            ? $"Réglage manuel ({FrequenceMoyenne} Hz)"
            : Designation;
    }
}