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
        /// Affiche la fenêtre d'aide, celle qui détaille chacun des champs à remplir.
        /// Purement visuel : ni le ViewModel ni l'enregistrement en cours ne sont touchés.
        /// </summary>
        private void OnOuvrirAide(object sender, RoutedEventArgs e)
        {
            var aide = new AideEnregistrementAppareilWindow { Owner = this };
            aide.ShowDialog();
        }
    }
}
