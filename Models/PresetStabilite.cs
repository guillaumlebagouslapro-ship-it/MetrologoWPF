using System.Collections.Generic;

namespace Metrologo.Models
{
    /// <summary>
    /// Un preset de balayage de stabilité : une série de temps de porte que l'orchestrator enchaîne
    /// l'un après l'autre (une feuille Excel par gate). Les libellés sont sur l'échelle canonique
    /// ("10 ms", "100 ms", "1 s"...) et ne sont traduits en indices qu'au moment de la mesure, pour
    /// rester indépendants des slots internes.
    /// </summary>
    public class PresetStabilite
    {
        public string Nom { get; set; } = string.Empty;

        /// <summary>Les libellés des gates à balayer, par ex. "10 ms", "100 ms", "1 s".</summary>
        public List<string> GatesLibelles { get; set; } = new();
    }
}
