using System.Diagnostics;

namespace Metrologo.Services
{
    /// <summary>
    /// Profileur léger : accumule des marqueurs (timestamp + label) en mémoire, à écrire dans un
    /// fichier texte après la mesure. Aucun I/O pendant la boucle GPIB (~10 ns par Mark).
    /// Format (temps en ms depuis la création de la session) :
    /// <code>
    /// Profiling FI=987654 | Type=Frequence | NbMesures=30 | Date=05/05/2026 10:00:00
    /// Total : 30450 ms
    ///
    /// Time(ms)  Delta(ms)  Étape
    /// ----------------------------------------------------------
    ///        0          0  Début mesure
    ///       12         12  Préparation Excel
    ///     1245       1233  Mesure 1: 999.99 Hz
    ///     2255       1010  Mesure 2: 999.98 Hz
    ///     ...
    ///    30450        250  Sauvegarde finale
    /// </code>
    /// </summary>
    public sealed class ProfilerSession
    {
        private readonly Stopwatch _sw = Stopwatch.StartNew();
        private readonly List<(long ts, string label)> _events = new();

        /// <summary>Ajoute un marqueur (~10 ns : Stopwatch.ElapsedMilliseconds + List.Add).</summary>
        public void Mark(string label)
        {
            _events.Add((_sw.ElapsedMilliseconds, label));
        }

        /// <summary>Nombre de marqueurs.</summary>
        public int Count => _events.Count;

    }
}
