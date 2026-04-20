using Metrologo.Models;
using Metrologo.ViewModels;
using System.Windows;

namespace Metrologo.Views
{
    public partial class LoginWindow : Window
    {
        public LoginWindow()
        {
            InitializeComponent();
        }

        public LoginWindow(object viewModel) : this()
        {
            DataContext = viewModel;
            
            // ModePosteCourant = null; (s'il existe dans le contexte)
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

            MessageBox.Show("Contexte de connexion invalide.", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}