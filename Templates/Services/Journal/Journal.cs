using System.Threading.Tasks;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Façade statique pour journaliser depuis n'importe où dans l'application.
    /// À initialiser avec <see cref="Configurer"/> au démarrage.
    /// </summary>
    public static class Journal
    {
        private static IJournalService? _service;

        public static IJournalService? Service => _service;
        public static string? SessionId => _service?.SessionActuelleId;
        public static string? Utilisateur => _service?.UtilisateurActuel;

        public static void Configurer(IJournalService service) => _service = service;

        public static Task DemarrerSessionAsync(string utilisateur)
            => _service?.DemarrerSessionAsync(utilisateur) ?? Task.CompletedTask;

        public static Task TerminerSessionAsync()
            => _service?.TerminerSessionAsync() ?? Task.CompletedTask;

        public static Task DefinirPosteAsync(string poste)
            => _service?.DefinirPosteAsync(poste) ?? Task.CompletedTask;

        public static Task NettoyerSessionsZombiesAsync()
            => _service?.NettoyerSessionsZombiesAsync() ?? Task.CompletedTask;

        public static void Info(CategorieLog cat, string action, string message, object? details = null)
            => _service?.Log(cat, action, message, details, SeveriteLog.Info);

        public static void Warn(CategorieLog cat, string action, string message, object? details = null)
            => _service?.Log(cat, action, message, details, SeveriteLog.Avertissement);

        public static void Erreur(CategorieLog cat, string action, string message, object? details = null)
            => _service?.Log(cat, action, message, details, SeveriteLog.Erreur);
    }
}
