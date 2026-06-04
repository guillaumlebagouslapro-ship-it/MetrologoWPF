using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace Metrologo.Views
{
    /// <summary>
    /// Petite notification non-bloquante (coin haut-droit) qui apparaît quelques secondes puis
    /// disparaît seule. Sert à signaler les changements admin propagés depuis un autre poste.
    /// N'active pas la fenêtre (ShowActivated=False) → ne vole pas le focus à une mesure en cours.
    /// </summary>
    public partial class ToastNotification : Window
    {
        private static readonly List<ToastNotification> _ouverts = new();
        private const double Marge = 16;
        private const double HauteurSlot = 96;   // espacement vertical entre toasts empilés

        private readonly DispatcherTimer _timer;

        private ToastNotification(string titre, string message, int dureeSecondes)
        {
            InitializeComponent();
            if (!string.IsNullOrWhiteSpace(titre)) TxtTitre.Text = titre;
            TxtMessage.Text = message;

            Loaded += (_, __) => Positionner();

            _timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(dureeSecondes) };
            _timer.Tick += (_, __) => Fermer();
            _timer.Start();

            Closed += (_, __) => { _ouverts.Remove(this); };
        }

        /// <summary>Affiche un toast sur le thread UI. Sûr à appeler depuis n'importe où.</summary>
        public static void Afficher(string titre, string message, int dureeSecondes = 6)
        {
            var disp = Application.Current?.Dispatcher;
            if (disp == null) return;
            disp.Invoke(() =>
            {
                var toast = new ToastNotification(titre, message, dureeSecondes);
                _ouverts.Add(toast);
                toast.Show();
            });
        }

        private void Positionner()
        {
            var zone = SystemParameters.WorkArea;
            Left = zone.Right - Width - Marge;
            int index = _ouverts.IndexOf(this);
            if (index < 0) index = _ouverts.Count - 1;
            Top = zone.Top + Marge + index * HauteurSlot;

            // Apparition en fondu.
            BeginAnimation(OpacityProperty,
                new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(180))));
        }

        private void Fermer()
        {
            _timer.Stop();
            var fade = new DoubleAnimation(Opacity, 0, new Duration(TimeSpan.FromMilliseconds(220)));
            fade.Completed += (_, __) => { try { Close(); } catch { } };
            BeginAnimation(OpacityProperty, fade);
        }
    }
}
