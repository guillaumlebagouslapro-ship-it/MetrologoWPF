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
        /// RadioButton « Catalogue » : coupe le mode manuel. L'activation est gérée par le
        /// binding TwoWay sur ModeManuel du radio « Réglage manuel ».
        /// </summary>
        private void OnModeListeChecked(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null) ViewModel.ModeManuel = false;
        }
    }
}
