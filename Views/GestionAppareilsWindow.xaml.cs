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
        /// Affiche la fiche d'aide (la même que celle d'EnregistrementAppareilWindow).
        /// C'est une fenêtre modale purement visuelle, elle ne modifie rien dans le catalogue.
        /// </summary>
        private void OnOuvrirAide(object sender, RoutedEventArgs e)
        {
            var aide = new AideEnregistrementAppareilWindow { Owner = this };
            aide.ShowDialog();
        }
    }
}
