using Metrologo;
using Metrologo.Services;
using Metrologo.ViewModels;
using Metrologo.Views;
using System.Windows;

namespace MetrologoWPF
{
    public partial class App : Application
    {
        protected override async void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // 1. Initialise la base de données (crée le fichier et les tables si besoin)
            await DatabaseInitializer.InitialiserAsync();

            // 2. Prépare et affiche la fenêtre de connexion
            var authService = new AuthService();
            var loginVM = new LoginViewModel(authService);
            var loginWin = new LoginWindow(loginVM);

            // 3. Si la connexion réussit
            if (loginWin.ShowDialog() == true)
            {
                // On passe l'utilisateur connecté au MainViewModel
                var mainVM = new MainViewModel
                {
                    UtilisateurConnecte = loginVM.UtilisateurSession
                };

                // On ouvre la fenêtre principale
                var mainWin = new MainWindow
                {
                    DataContext = mainVM
                };

                Application.Current.MainWindow = mainWin;
                mainWin.Show();
            }
            else
            {
                // Si l'utilisateur ferme la fenêtre de login, on quitte l'appli
                Shutdown();
            }
        }
    }
}