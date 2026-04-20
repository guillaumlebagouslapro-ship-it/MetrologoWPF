using Metrologo;
using Metrologo.Services;
using Metrologo.ViewModels;
using Metrologo.Views;
using System.Windows;
using System.Threading.Tasks;

namespace MetrologoWPF
{
    public partial class App : Application
    {
        protected override async void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;

            // Initialize DB asynchrone pour ne pas tuer le Thread UI (évite de freeze ou de crash la fenêtre noire)
            await DatabaseInitializer.InitialiserAsync();

            var authService = new AuthService();
            var loginVM = new LoginViewModel(authService);
            var loginWin = new Metrologo.Views.LoginWindow(loginVM);

            loginVM.CloseAction = ok =>
            {
                // Si la fenêtre de login est encore ouverte, on applique la réponse
                if (loginWin.IsVisible)
                {
                    loginWin.DialogResult = ok;
                    loginWin.Close();
                }
            };

            // L'appel bloquant ShowDialog relancera la UI après initialisation DB
            bool? result = loginWin.ShowDialog();

            if (result == true)
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
                Application.Current.ShutdownMode = ShutdownMode.OnMainWindowClose;
                mainWin.Show();
            }
            else
            {
                Application.Current.Shutdown();
            }
        }
    }
}