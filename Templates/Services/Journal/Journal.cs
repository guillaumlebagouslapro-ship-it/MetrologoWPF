using System.Threading.Tasks;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Point d'entrée statique pour journaliser depuis n'importe où dans l'application.
    /// À initialiser au démarrage via <see cref="Configurer"/>.
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
            => Router(cat, action, message, details, SeveriteLog.Info);

        public static void Warn(CategorieLog cat, string action, string message, object? details = null)
            => Router(cat, action, message, details, SeveriteLog.Avertissement);

        public static void Erreur(CategorieLog cat, string action, string message, object? details = null)
            => Router(cat, action, message, details, SeveriteLog.Erreur);

        /// <summary>
        /// Écrit dans le journal via le service, et au passage recopie dans le journal d'AUDIT
        /// administrateur les actions de configuration (catégories Administration/Rubidium dont le
        /// code est retenu — voir <see cref="JournalAdminService"/>). Les simples consultations
        /// (OUVERTURE_*) ne sont pas dans la liste blanche, donc on les laisse de côté.
        /// </summary>
        private static void Router(CategorieLog cat, string action, string message, object? details, SeveriteLog sev)
        {
            _service?.Log(cat, action, message, details, sev);

            if ((cat == CategorieLog.Administration || cat == CategorieLog.Rubidium)
                && JournalAdminService.EstActionAudit(action))
            {
                JournalAdminService.Ecrire(action, message, Utilisateur);
            }
        }
    }
}
