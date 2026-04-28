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

        /// <summary>Met à jour le poste de la session courante (Baie / Paillasse).</summary>
        Task DefinirPosteAsync(string poste);

        /// <summary>Ferme les sessions laissées ouvertes par un arrêt non propre.</summary>
        Task NettoyerSessionsZombiesAsync();

        Task LogAsync(CategorieLog categorie, string action, string message, object? details = null, SeveriteLog severite = SeveriteLog.Info);

        /// <summary>Fire-and-forget — n'attend pas l'écriture DB.</summary>
        void Log(CategorieLog categorie, string action, string message, object? details = null, SeveriteLog severite = SeveriteLog.Info);

        Task<List<SessionJournal>> ChargerSessionsAsync(FiltreJournal? filtre = null);
        Task<List<string>> ChargerListeUtilisateursAsync();
    }
}
