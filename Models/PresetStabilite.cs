using System.Collections.Generic;

namespace Metrologo.Models
{
    /// <summary>
    /// Preset de balayage de stabilité : un ensemble de temps de porte que l'orchestrator
    /// va parcourir séquentiellement, créant une feuille Excel par gate. Les libellés
    /// suivent l'échelle canonique (« 10 ms », « 100 ms », « 1 s », « 10 s »…) et sont
    /// résolus en indices au moment de la mesure pour rester découplés des slots internes.
    /// </summary>
    public class PresetStabilite
    {
        public string Nom { get; set; } = string.Empty;

        /// <summary>Libellés des gates à balayer (ex: « 10 ms », « 100 ms », « 1 s »).</summary>
        public List<string> GatesLibelles { get; set; } = new();
    }
}
