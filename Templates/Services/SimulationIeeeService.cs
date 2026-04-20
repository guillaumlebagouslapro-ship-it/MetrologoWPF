using System;
using System.Threading;
using System.Threading.Tasks;
using Metrologo.Models;

namespace Metrologo.Services
{
    public class SimulationIeeeService : IIeeeService
    {
        private readonly Random _random = new Random();

        // Gate time en secondes, indexé par GateIndex (0..12)
        private static readonly double[] GateSeconds =
        {
            0.010, 0.020, 0.050, 0.100, 0.200, 0.500,
            1.0, 2.0, 5.0, 10.0, 20.0, 50.0, 100.0
        };

        public async Task<bool> InitialiserAsync()
        {
            await Task.Delay(500);
            return true;
        }

        public async Task<double> LireMesureAsync(Mesure mesure, CancellationToken ct = default)
        {
            var gateIndex = mesure.GateIndex;
            // Sentinelles procédures auto (-3, -2, -1) → défaut 1s pour la simulation
            double gateS = (gateIndex >= 0 && gateIndex < GateSeconds.Length)
                ? GateSeconds[gateIndex]
                : 1.0;

            // Simule le temps d'intégration réel (borné à 2s max pour ne pas bloquer l'UX)
            var delayMs = (int)Math.Min(gateS * 1000, 2000);
            await Task.Delay(delayMs, ct);

            // Fréquence nominale ~ 10 MHz
            double fNominale = mesure.FNominale > 0 ? mesure.FNominale : 10_000_000.0;

            // Bruit blanc ~ 1/√(gate) × facteur instrument
            double facteurInstrument = mesure.Frequencemetre switch
            {
                TypeAppareilIEEE.Stanford => 0.04,  // SR620 très bon
                TypeAppareilIEEE.Racal    => 0.08,
                TypeAppareilIEEE.EIP      => 0.10,
                _ => 0.06
            };
            double sigma = facteurInstrument / Math.Sqrt(gateS);

            // Approximation gaussienne via Box-Muller
            double u1 = 1.0 - _random.NextDouble();
            double u2 = 1.0 - _random.NextDouble();
            double gauss = Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Cos(2.0 * Math.PI * u2);
            double bruit = sigma * gauss;

            // Léger biais systématique selon l'instrument (ppb)
            double biaisPpb = mesure.Frequencemetre switch
            {
                TypeAppareilIEEE.Stanford => 0.0,
                TypeAppareilIEEE.Racal    => -0.02,
                TypeAppareilIEEE.EIP      => 0.05,
                _ => 0.0
            };
            double biais = fNominale * biaisPpb * 1e-9;

            return fNominale + biais + bruit;
        }
    }
}
