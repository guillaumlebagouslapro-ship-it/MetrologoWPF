using System.Threading.Tasks;
using Metrologo.Models;

namespace Metrologo.Services
{
    public interface IIeeeService
    {
        // Pour tester si le bus IEEE répond
        Task<bool> InitialiserAsync();

        // Pour lire une valeur depuis un fréquencemètre
        Task<double> LireMesureAsync(AppareilIEEE appareil);
    }
}