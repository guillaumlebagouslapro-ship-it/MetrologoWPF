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

            await DatabaseInitializer.InitialiserAsync();

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

                MainWindow = mainWin;
                mainWin.Show();
            }
            else
            {
                Shutdown();
            }
        }
    }
}