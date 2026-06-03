namespace Metrologo.Models
{
    public class Rubidium
    {
        public int Id { get; set; }
        public string Designation { get; set; } = string.Empty;
        public double FrequenceMoyenne { get; set; }
        public bool AvecGPS { get; set; }

        /// <summary>
        /// Vrai si la fréquence a été saisie manuellement par l'admin (= aucun rubidium
        /// du catalogue n'a été sélectionné). Dans ce cas, la valeur de
        /// <see cref="FrequenceMoyenne"/> vient directement de la saisie utilisateur.
        /// </summary>
        public bool EstReglageManuel { get; set; }

        // Affiche directement le nom choisi dans le catalogue (réglage manuel à part).
        public string NomAffichage => EstReglageManuel
            ? $"Réglage manuel ({FrequenceMoyenne} Hz)"
            : Designation;

        /// <summary>
        /// Libellé écrit dans la feuille de mesure (zone <c>ZNRubidium</c>). On affiche
        /// directement la <see cref="Designation"/> choisie dans le catalogue (le nom est libre :
        /// y inclure une mention de raccordement — « GPS » etc. — si besoin). Aucun suffixe
        /// automatique n'est ajouté. Pour un réglage manuel : « Réglage manuel (xxx Hz) ».
        /// </summary>
        public string DesignationAvecRaccord => EstReglageManuel
            ? $"Réglage manuel ({FrequenceMoyenne} Hz)"
            : Designation;
    }
}