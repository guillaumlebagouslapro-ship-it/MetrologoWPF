using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Threading;
using Metrologo.Views;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Garde un œil sur le journal d'audit administrateur partagé et fait apparaître un toast
    /// (sans bloquer) sur CE poste dès qu'un AUTRE poste a modifié la configuration : rubidium,
    /// chemins, modules d'incertitude, catalogue, utilisateurs… On relit le fichier de façon
    /// incrémentale toutes les 12 s. Démarré au lancement de l'app, sur le thread UI.
    /// </summary>
    public static class NotificationsAdminWatcher
    {
        private static DispatcherTimer? _timer;
        private static long _position;
        private static readonly string _machine = Environment.MachineName;

        /// <summary>
        /// Déclenché (sur le thread UI) à chaque lot de changements admin venus d'un AUTRE poste.
        /// La barre de navigation s'en sert pour afficher un indicateur persistant (triangle ⚠).
        /// </summary>
        public static event Action<IReadOnlyList<EntreeJournalAdmin>>? ChangementsRecus;

        public static void Demarrer()
        {
            if (_timer != null) return;

            // Point de départ : on se cale sur la fin du fichier, pour ne pas notifier ce qui
            // s'est passé AVANT le lancement de ce poste, mais seulement ce qui arrivera ensuite.
            _position = JournalAdminService.Position();

            _timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(12) };
            _timer.Tick += (_, __) => Verifier();
            _timer.Start();
        }

        private static void Verifier()
        {
            List<EntreeJournalAdmin> nouvelles;
            try { nouvelles = JournalAdminService.LireDepuis(ref _position); }
            catch { return; }
            if (nouvelles.Count == 0) return;

            // On ne notifie que les CHANGEMENTS d'un AUTRE poste (pas les nôtres), et on laisse
            // de côté les simples accès admin (login) qui ne touchent pas à la configuration.
            var autres = nouvelles
                .Where(e => !string.Equals(e.Machine, _machine, StringComparison.OrdinalIgnoreCase))
                .Where(e => !e.Action.StartsWith("ACCES_ADMIN", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (autres.Count == 0) return;

            // Indicateur qui reste affiché dans le bandeau (en plus du toast passager).
            try { ChangementsRecus?.Invoke(autres); } catch { }

            if (autres.Count == 1)
            {
                var e = autres[0];
                string msg = e.ActionLisible
                    + (string.IsNullOrWhiteSpace(e.Detail) ? string.Empty : $" — {e.Detail}")
                    + (string.IsNullOrWhiteSpace(e.Utilisateur) ? string.Empty : $"  (par {e.Utilisateur})");
                ToastNotification.Afficher("Changement pris en compte", msg);
            }
            else
            {
                ToastNotification.Afficher("Changements pris en compte",
                    $"{autres.Count} modifications administrateur appliquées (voir Journal administrateur).");
            }
        }
    }
}
