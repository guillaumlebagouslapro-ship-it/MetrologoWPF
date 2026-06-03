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

        // Une petite propriété pratique qui formate le nom tout seul
        public string NomAffichage => EstReglageManuel
            ? $"Réglage manuel ({FrequenceMoyenne:N2} Hz)"
            : (AvecGPS ? $"GPS + {Designation}" : Designation);

        /// <summary>
        /// Libellé écrit dans la feuille de mesure (zone <c>ZNRubidium</c>) : désignation +
        /// mode de raccordement réellement sélectionné en admin. Trois cas :
        ///   - réglage manuel  → "Réglage manuel (xxx Hz)" (pas de raccord Allouis/GPS) ;
        ///   - GPS             → "<désignation> (raccord GPS)" ;
        ///   - sinon           → "<désignation> (raccord Allouis)".
        /// Évite d'afficher « raccord Allouis » par défaut quand le rubidium est en manuel.
        /// </summary>
        public string DesignationAvecRaccord => EstReglageManuel
            ? $"Réglage manuel ({FrequenceMoyenne:N2} Hz)"
            : Designation + (AvecGPS ? " (raccord GPS)" : " (raccord Allouis)");
    }
}