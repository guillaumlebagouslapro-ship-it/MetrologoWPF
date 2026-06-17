using System;
using System.Threading.Tasks;
using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    /// <summary>
    /// Boîte de dialogue bloquante qu'on montre au démarrage quand le partage réseau
    /// serveur (M:\exe_spe\Data_Metrologo) ne répond pas. L'utilisateur a deux choix :
    ///   • « Rafraîchir » — retente l'accès et laisse la fenêtre ouverte tant que le
    ///     partage ne répond toujours pas ; dès qu'il répond, DialogResult = true.
    ///   • « Fermer l'application » — DialogResult = false, et l'appelant coupe l'app.
    /// </summary>
    public partial class ReseauIndisponibleDialog : FluentWindow
    {
        private readonly Func<bool> _testerAcces;
        private int _tentatives;

        /// <param name="cheminPartage">Chemin du partage affiché à l'utilisateur.</param>
        /// <param name="testerAcces">Teste l'accès au partage et renvoie true s'il est joignable.
        /// Tourne en dehors du thread UI, car sur un lecteur réseau mort il peut rester bloqué
        /// plusieurs secondes.</param>
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
                // On passe par Task.Run parce qu'un lecteur réseau démonté peut figer
                // plusieurs secondes au premier accès : autant ne pas geler l'UI pendant
                // ce temps-là.
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
