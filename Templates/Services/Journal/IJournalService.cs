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

        /// <summary>Tire et oublie : on n'attend pas que l'écriture soit faite.</summary>
        void Log(CategorieLog categorie, string action, string message, object? details = null, SeveriteLog severite = SeveriteLog.Info);

        /// <summary>Toutes les entrées (y compris celles écrites hors session, au démarrage par
        /// ex.), les plus récentes d'abord, avec utilisateur et machine renseignés.</summary>
        Task<List<LogEntry>> ChargerEntreesAsync(FiltreJournal? filtre = null);
    }
}
