using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace Metrologo.Services.Ieee
{
    /// <summary>
    /// Pilote IEEE simulé, aucun matériel (dev / tests / formation). Mémorise la dernière commande
    /// par adresse ; si elle ressemble à une requête de mesure (meas?, RS, RE), renvoie ~10 MHz
    /// bruité avec un sigma dépendant de la gate. Le MAV est toujours prêt après un court délai.
    /// </summary>
    public class SimulationIeeeDriver : IIeeeDriver
    {
        private readonly Random _random = new Random();
        private readonly Dictionary<int, string> _derniereCommande = new();

        // temps de gate (s) extrait grossièrement de la dernière commande gate reçue
        private readonly Dictionary<int, double> _gateSecondes = new();

        public Task SendInterfaceClearAsync(CancellationToken ct = default)
            => Task.CompletedTask;

        public Task EcrireAsync(int adresse, string commande, int writeTerm, CancellationToken ct = default)
        {
            var cmd = commande ?? string.Empty;
            _derniereCommande[adresse] = cmd;
            DetecterGate(adresse, cmd);
            return Task.CompletedTask;
        }

        public async Task<string> LireAsync(int adresse, int readTerm, CancellationToken ct = default)
        {
            await Task.Delay(20, ct);

            var cmd = _derniereCommande.GetValueOrDefault(adresse, string.Empty);
            if (!EstCommandeMesure(cmd)) return string.Empty;

            double gateS = _gateSecondes.GetValueOrDefault(adresse, 1.0);
            double sigma = 0.05 / Math.Sqrt(gateS);
            double bruit = sigma * GaussienneReduite();
            double valeur = 10_000_000.0 + bruit;

            return valeur.ToString("F6", CultureInfo.InvariantCulture);
        }

        public async Task<byte> LireStatusByteAsync(int adresse, CancellationToken ct = default)
        {
            // court délai d'acquisition puis MAV armé
            await Task.Delay(10, ct);
            return 0x10;  // MAV set
        }

        public Task DefinirRemoteLocalAsync(int adresse, bool remote, CancellationToken ct = default)
            => Task.CompletedTask;

        public void ReinitialiserSessions()
        {
            // pas de session persistante en simulation
            _derniereCommande.Clear();
            _gateSecondes.Clear();
        }

        public void AborterToutesSessions()
        {
            // rien à débloquer : le LireAsync simulé est court et déjà annulable via le token
        }

        public void DefinirTimeout(int adresse, int timeoutMs)
        {
            // no-op en simulation
        }

        // --------------- Internes ---------------

        private static bool EstCommandeMesure(string cmd)
        {
            if (string.IsNullOrEmpty(cmd)) return false;
            var c = cmd.ToLowerInvariant();
            return c.StartsWith("meas") || c.StartsWith("rs") || c.StartsWith("re");
        }

        // reconnaît grossièrement les notations type 1E0, 1E1, GA1E-2 pour fixer la gate simulée
        private void DetecterGate(int adresse, string cmd)
        {
            if (string.IsNullOrEmpty(cmd)) return;
            int idx = cmd.IndexOf('E', StringComparison.OrdinalIgnoreCase);
            if (idx < 1 || idx >= cmd.Length - 1) return;

            // mantisse juste avant 'E', exposant (signé) juste après
            int mStart = idx - 1;
            while (mStart > 0 && char.IsDigit(cmd[mStart - 1])) mStart--;
            string mantisse = cmd.Substring(mStart, idx - mStart);

            int eEnd = idx + 1;
            if (eEnd < cmd.Length && (cmd[eEnd] == '+' || cmd[eEnd] == '-')) eEnd++;
            while (eEnd < cmd.Length && char.IsDigit(cmd[eEnd])) eEnd++;
            string exposant = cmd.Substring(idx + 1, eEnd - idx - 1);

            if (!double.TryParse(mantisse, NumberStyles.Float, CultureInfo.InvariantCulture, out var m)) return;
            if (!int.TryParse(exposant, NumberStyles.Integer, CultureInfo.InvariantCulture, out var e)) return;

            double seconds = m * Math.Pow(10, e);
            if (seconds > 0 && seconds <= 100) _gateSecondes[adresse] = seconds;
        }

        private double GaussienneReduite()
        {
            double u1 = 1.0 - _random.NextDouble();
            double u2 = 1.0 - _random.NextDouble();
            return Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Cos(2.0 * Math.PI * u2);
        }
    }
}
