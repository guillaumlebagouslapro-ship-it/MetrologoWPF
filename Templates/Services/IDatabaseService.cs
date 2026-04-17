using Metrologo.Models;
using System.Threading.Tasks;

namespace Metrologo.Services
{
    public interface IDatabaseService
    {
        Task<bool> TesterConnexionAsync();
        // La nouvelle méthode (le "?" veut dire qu'elle peut renvoyer null si la base est vide)
        Task<Rubidium?> GetRubidiumActifAsync();
    }
}