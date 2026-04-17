using Metrologo.ViewModels;
using System.Windows;

namespace Metrologo.Views
{
    public partial class LoginWindow : Window
    {
        public LoginWindow(LoginViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;

            // Permet au ViewModel de fermer cette fenêtre
            viewModel.CloseAction = (result) =>
            {
                this.DialogResult = result;
                this.Close();
            };
        }

        // On gère le clic ici pour récupérer le mot de passe de la PasswordBox de façon sécurisée
        private async void BtnConnexion_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is LoginViewModel vm)
            {
                await vm.SeConnecterAsync(txtMotDePasse.Password);
            }
        }
    }
}