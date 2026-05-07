using System.Windows;
using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class GestionAppareilsWindow : FluentWindow
    {
        public GestionAppareilsWindow(GestionAppareilsViewModel vm)
        {
            InitializeComponent();
            DataContext = vm;
        }

        /// <summary>
        /// Ouvre la fiche d'aide (même contenu que pour EnregistrementAppareilWindow).
        /// Pure UI, modale, ne touche rien du catalogue.
        /// </summary>
        private void OnOuvrirAide(object sender, RoutedEventArgs e)
        {
            var aide = new AideEnregistrementAppareilWindow { Owner = this };
            aide.ShowDialog();
        }
    }
}
