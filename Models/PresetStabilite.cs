using System.Collections.Generic;

namespace Metrologo.Models
{
    /// <summary>
    /// Preset de balayage de stabilité : des temps de porte parcourus séquentiellement par
    /// l'orchestrator (une feuille Excel par gate). Libellés sur l'échelle canonique ("10 ms",
    /// "100 ms", "1 s"...), résolus en indices à la mesure pour rester découplés des slots internes.
    /// </summary>
    public class PresetStabilite
    {
        public string Nom { get; set; } = string.Empty;

        /// <summary>Libellés des gates à balayer (ex: "10 ms", "100 ms", "1 s").</summary>
        public List<string> GatesLibelles { get; set; } = new();
    }
}
