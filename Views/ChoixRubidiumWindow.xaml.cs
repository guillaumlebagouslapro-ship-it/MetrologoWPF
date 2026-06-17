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
        /// Handler du RadioButton « Catalogue ». Quand l'utilisateur revient sur le catalogue,
        /// on coupe le mode manuel. L'activation, elle, est gérée par le binding TwoWay sur
        /// ModeManuel côté radio « Réglage manuel » ; ici on s'occupe juste du retour en arrière.
        /// </summary>
        private void OnModeListeChecked(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.ModeManuel = false;
        }
    }
}
