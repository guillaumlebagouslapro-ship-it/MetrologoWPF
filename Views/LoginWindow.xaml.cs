using Metrologo.ViewModels;
using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class LoginWindow : FluentWindow
    {
        public LoginWindow()
        {
            InitializeComponent();
        }

        public LoginWindow(object viewModel) : this()
        {
            DataContext = viewModel;
        }

        private async void BtnConnexion_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is LoginViewModel vm)
            {
                await vm.SeConnecterAsync(txtMotDePasse.Password);

                if (vm.UtilisateurSession != null && !DialogResult.HasValue)
                {
                    DialogResult = true;
                    Close();
                }

                return;
            }

            System.Windows.MessageBox.Show("Contexte de connexion invalide.", "Erreur", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }
}
