using System.Threading;
using System.Threading.Tasks;
using Metrologo.Models;

namespace Metrologo.Services
{
    public interface IIeeeService
    {
        Task<bool> InitialiserAsync();

        /// <summary>
        /// Lit une valeur de fréquence en respectant le gate time configuré dans la Mesure.
        /// </summary>
        Task<double> LireMesureAsync(Mesure mesure, CancellationToken ct = default);
    }
}
