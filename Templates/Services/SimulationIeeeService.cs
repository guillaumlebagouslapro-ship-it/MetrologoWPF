using System;
using System.Threading.Tasks;
using Metrologo.Models;

namespace Metrologo.Services
{
    public class SimulationIeeeService : IIeeeService
    {
        private readonly Random _random = new Random();

        public async Task<bool> InitialiserAsync()
        {
            // On simule un petit temps de chargement du bus IEEE
            await Task.Delay(500);
            return true;
        }

        public async Task<double> LireMesureAsync(AppareilIEEE appareil)
        {
            // On simule le temps de la porte (Gate) du fréquencemètre
            await Task.Delay(250);

            // Fausse fréquence autour de 10 MHz (10 000 000 Hz) avec un micro bruit aléatoire
            double frequenceDeBase = 10000000.0;
            double bruit = (_random.NextDouble() - 0.5) * 0.05;

            return frequenceDeBase + bruit;
        }
    }
}