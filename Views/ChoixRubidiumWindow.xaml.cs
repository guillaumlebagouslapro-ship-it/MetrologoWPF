using System.Windows;
using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class ChoixRubidiumWindow : FluentWindow
    {
        public ChoixRubidiumViewModel ViewModel { get; }

        public ChoixRubidiumWindow()
        {
            InitializeComponent();
            ViewModel = new ChoixRubidiumViewModel();
            DataContext = ViewModel;
            ViewModel.CloseAction = result => { DialogResult = result; Close(); };
        }

        /// <summary>
        /// Handler du RadioButton « Catalogue » : désactive le mode manuel quand
        /// l'utilisateur revient sur le catalogue. Le binding TwoWay sur ModeManuel
        /// (côté radio « Réglage manuel ») gère l'activation ; ici on assure le retour.
        /// </summary>
        private void OnModeListeChecked(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.ModeManuel = false;
        }
    }
}
