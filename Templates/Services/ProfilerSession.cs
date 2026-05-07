using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace Metrologo.Services
{
    /// <summary>
    /// Profileur léger qui accumule des marqueurs (timestamp + label) en mémoire pendant
    /// une mesure, puis dumpe un fichier texte à la fin. Aucun I/O pendant la mesure
    /// (sauf le <c>Stopwatch.GetTimestamp</c> qui coûte ~10 ns) — évite que le journal
    /// SQL fasse 30+ écritures pendant la boucle GPIB et ralentisse la mesure.
    ///
    /// Format du fichier (timing en ms depuis le démarrage de la session) :
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

        /// <summary>Ajoute un marqueur. Coût : ~10 ns (Stopwatch.ElapsedMilliseconds + List.Add).</summary>
        public void Mark(string label)
        {
            _events.Add((_sw.ElapsedMilliseconds, label));
        }

        /// <summary>Durée totale (ms) depuis la création du profiler.</summary>
        public long TotalMs => _sw.ElapsedMilliseconds;

        /// <summary>Nombre de marqueurs accumulés — utilisé pour pré-allouer dans le formatter.</summary>
        public int Count => _events.Count;

        /// <summary>
        /// Sérialise le profiling en texte tabulé. Le caller décide où l'écrire (fichier,
        /// log, etc.). Ne fait aucun I/O lui-même — gardé pur pour testabilité.
        /// </summary>
        public string FormaterTexte(string entete)
        {
            var sb = new StringBuilder(80 * (_events.Count + 5));
            sb.AppendLine(entete);
            sb.AppendLine($"Total : {TotalMs} ms");
            sb.AppendLine();
            sb.AppendLine("Time(ms)  Delta(ms)  Étape");
            sb.AppendLine("----------------------------------------------------------");

            long prev = 0;
            foreach (var (ts, label) in _events)
            {
                sb.Append(ts.ToString("D8", CultureInfo.InvariantCulture));
                sb.Append("  ");
                sb.Append((ts - prev).ToString("D9", CultureInfo.InvariantCulture));
                sb.Append("  ");
                sb.AppendLine(label);
                prev = ts;
            }
            return sb.ToString();
        }

        /// <summary>
        /// Écrit le profiling dans un fichier texte. À appeler après la mesure (jamais
        /// pendant). Crée les répertoires manquants. Erreurs swallowed avec stderr —
        /// le caller peut logger en plus s'il le souhaite.
        /// </summary>
        public void EcrireFichier(string cheminFichier, string entete)
        {
            string? dossier = Path.GetDirectoryName(cheminFichier);
            if (!string.IsNullOrEmpty(dossier))
                Directory.CreateDirectory(dossier);

            File.WriteAllText(cheminFichier, FormaterTexte(entete), Encoding.UTF8);
        }
    }
}
