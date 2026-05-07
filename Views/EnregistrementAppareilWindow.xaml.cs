using System.Windows;
using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class EnregistrementAppareilWindow : FluentWindow
    {
        public EnregistrementAppareilViewModel ViewModel { get; }

        public EnregistrementAppareilWindow(EnregistrementAppareilViewModel vm)
        {
            InitializeComponent();
            ViewModel = vm;
            DataContext = vm;
            vm.CloseAction = ok =>
            {
                if (IsVisible) { DialogResult = ok; Close(); }
            };
        }

        /// <summary>
        /// Ouvre la fenêtre d'aide qui explique chaque paramètre de saisie. Pure UI —
        /// aucun impact sur le ViewModel ni sur l'enregistrement en cours.
        /// </summary>
        private void OnOuvrirAide(object sender, RoutedEventArgs e)
        {
            var aide = new AideEnregistrementAppareilWindow { Owner = this };
            aide.ShowDialog();
        }
    }
}
