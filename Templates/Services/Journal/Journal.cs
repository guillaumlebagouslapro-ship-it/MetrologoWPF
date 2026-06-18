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

        /// <summary>
        /// Active les traces verbeuses (<see cref="Trace"/>) : log par mesure (envoi/réception
        /// GPIB, attente SRQ...). Désactivé par défaut — ces entrées, à raison de plusieurs par
        /// mesure × 30 mesures × N gates, gonflent vite le journal et ralentissent son ouverture,
        /// sans intérêt en routine. À mettre à true seulement pour diagnostiquer un appareil muet.
        /// </summary>
        public static bool Verbeux { get; set; }

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

        /// <summary>Trace de diagnostic fine : n'écrit QUE si <see cref="Verbeux"/> est actif.
        /// Pour les logs par-mesure du chemin chaud (GPIB), qui ne doivent pas grossir le journal
        /// en fonctionnement normal.</summary>
        public static void Trace(CategorieLog cat, string action, string message, object? details = null)
        {
            if (Verbeux) Router(cat, action, message, details, SeveriteLog.Info);
        }

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
