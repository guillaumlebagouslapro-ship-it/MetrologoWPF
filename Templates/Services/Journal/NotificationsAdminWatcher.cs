using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Threading;
using Metrologo.Views;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Surveille le journal d'audit administrateur partagé et affiche un toast non-bloquant sur
    /// CE poste quand un changement de configuration a été fait depuis un AUTRE poste (rubidium,
    /// chemins, modules d'incertitude, catalogue, utilisateurs…). Lecture incrémentale toutes les
    /// 12 s. Démarré au lancement de l'app (thread UI).
    /// </summary>
    public static class NotificationsAdminWatcher
    {
        private static DispatcherTimer? _timer;
        private static long _position;
        private static readonly string _machine = Environment.MachineName;

        /// <summary>
        /// Levé (sur le thread UI) à chaque lot de changements admin reçus depuis un AUTRE poste.
        /// Permet à la barre de navigation d'afficher un indicateur persistant (triangle ⚠).
        /// </summary>
        public static event Action<IReadOnlyList<EntreeJournalAdmin>>? ChangementsRecus;

        public static void Demarrer()
        {
            if (_timer != null) return;

            // Baseline : on part de la fin du fichier → on ne notifie pas pour ce qui s'est
            // passé AVANT le lancement de ce poste, seulement les changements à venir.
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

            // Ne notifie que les CHANGEMENTS faits par un AUTRE poste (pas soi-même), en
            // excluant les simples accès admin (login) qui ne modifient pas la configuration.
            var autres = nouvelles
                .Where(e => !string.Equals(e.Machine, _machine, StringComparison.OrdinalIgnoreCase))
                .Where(e => !e.Action.StartsWith("ACCES_ADMIN", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (autres.Count == 0) return;

            // Indicateur persistant dans le bandeau (en plus du toast éphémère).
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
