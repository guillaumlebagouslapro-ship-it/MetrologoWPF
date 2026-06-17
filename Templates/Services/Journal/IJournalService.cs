using System.Collections.Generic;
using System.Threading.Tasks;

namespace Metrologo.Services.Journal
{
    public interface IJournalService
    {
        string? SessionActuelleId { get; }
        string? UtilisateurActuel { get; }

        Task DemarrerSessionAsync(string utilisateur);
        Task TerminerSessionAsync();

        /// <summary>Renseigne le poste de la session en cours (Baie / Paillasse).</summary>
        Task DefinirPosteAsync(string poste);

        /// <summary>Referme les sessions restées ouvertes suite à un arrêt brutal.</summary>
        Task NettoyerSessionsZombiesAsync();

        Task LogAsync(CategorieLog categorie, string action, string message, object? details = null, SeveriteLog severite = SeveriteLog.Info);

        /// <summary>Tire et oublie : on n'attend pas que l'écriture soit faite.</summary>
        void Log(CategorieLog categorie, string action, string message, object? details = null, SeveriteLog severite = SeveriteLog.Info);

        Task<List<SessionJournal>> ChargerSessionsAsync(FiltreJournal? filtre = null);
        Task<List<string>> ChargerListeUtilisateursAsync();
    }
}
