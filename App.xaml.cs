using Metrologo;
using Metrologo.Services;
using Metrologo.ViewModels;
using Metrologo.Views;
using System.Windows;

namespace MetrologoWPF
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // 1. FORCER LA CRÉATION DE LA BDD AVANT TOUT LE RESTE (.Wait() est important ici)
            DatabaseInitializer.InitialiserAsync().Wait();

            // 2. Fenêtre de Login
            var authService = new AuthService();
            var loginVM = new LoginViewModel(authService);
            var loginWin = new LoginWindow(loginVM);

            if (loginWin.ShowDialog() == true)
            {
                var mainVM = new MainViewModel
                {
                    UtilisateurConnecte = loginVM.UtilisateurSession
                };

                var mainWin = new MainWindow
                {
                    DataContext = mainVM
                };

                Application.Current.MainWindow = mainWin;
                mainWin.Show();
            }
            else
            {
                Shutdown();
            }
        }
    }
}