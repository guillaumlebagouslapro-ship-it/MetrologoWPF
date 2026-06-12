using System;
using System.Threading.Tasks;
using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    /// <summary>
    /// Dialogue bloquant affiché au démarrage quand le partage réseau serveur
    /// (M:\exe_spe\Data_Metrologo) est inaccessible. Deux issues :
    ///   • « Rafraîchir » — relance le test d'accès (sans fermer la fenêtre tant
    ///     que le partage ne répond pas) ; dès qu'il répond, DialogResult = true.
    ///   • « Fermer l'application » — DialogResult = false, l'appelant arrête l'app.
    /// </summary>
    public partial class ReseauIndisponibleDialog : FluentWindow
    {
        private readonly Func<bool> _testerAcces;
        private int _tentatives;

        /// <param name="cheminPartage">Chemin du partage affiché à l'utilisateur.</param>
        /// <param name="testerAcces">Test d'accès au partage — retourne true si joignable.
        /// Exécuté hors thread UI (peut bloquer plusieurs secondes sur un lecteur réseau mort).</param>
        public ReseauIndisponibleDialog(string cheminPartage, Func<bool> testerAcces)
        {
            InitializeComponent();
            _testerAcces = testerAcces;
            RunChemin.Text = $"« {cheminPartage} »";
        }

        private async void OnRafraichir(object sender, RoutedEventArgs e)
        {
            BtnRafraichir.IsEnabled = false;
            BtnFermer.IsEnabled = false;
            BordureStatut.Visibility = Visibility.Visible;
            BordureStatut.Background = System.Windows.Media.Brushes.Transparent;
            IconeStatut.Symbol = SymbolRegular.ArrowSync24;
            IconeStatut.Foreground = (System.Windows.Media.Brush)FindResource("AsertiNavyBrush");
            TbStatut.Foreground = IconeStatut.Foreground;
            TbStatut.Text = "Vérification de l'accès au partage en cours…";

            bool accessible;
            try
            {
                // Task.Run : un lecteur réseau démonté peut bloquer plusieurs secondes
                // sur le premier accès — on ne gèle pas l'UI pendant ce temps.
                accessible = await Task.Run(_testerAcces);
            }
            catch
            {
                accessible = false;
            }

            if (accessible)
            {
                DialogResult = true;
                Close();
                return;
            }

            _tentatives++;
            BordureStatut.Background = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(0xFE, 0xE2, 0xE2));
            IconeStatut.Symbol = SymbolRegular.ErrorCircle24;
            IconeStatut.Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(0xDC, 0x26, 0x26));
            TbStatut.Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(0x99, 0x1B, 0x1B));
            TbStatut.Text = _tentatives == 1
                ? "Le partage est toujours inaccessible. Vérifie la connexion réseau puis réessaie."
                : $"Le partage est toujours inaccessible ({_tentatives} tentatives). "
                  + "Vérifie la connexion réseau puis réessaie.";
            BtnRafraichir.IsEnabled = true;
            BtnFermer.IsEnabled = true;
        }

        private void OnFermer(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
