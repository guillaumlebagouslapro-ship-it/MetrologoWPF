using Metrologo;
using Metrologo.Services;
using Metrologo.Services.Journal;
using Metrologo.ViewModels;
using Metrologo.Views;
using System.Windows;
using System.Threading.Tasks;

namespace Metrologo
{
    public partial class App : Application
    {
        protected override async void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;

            await DatabaseInitializer.InitialiserAsync();

            // Initialisation du journal (SQLite local)
            Journal.Configurer(new SqliteJournalService());

            var authService = new AuthService();
            var loginVM = new LoginViewModel(authService);
            var loginWin = new Metrologo.Views.LoginWindow(loginVM);

            loginVM.CloseAction = ok =>
            {
                if (loginWin.IsVisible)
                {
                    loginWin.DialogResult = ok;
                    loginWin.Close();
                }
            };

            bool? result = loginWin.ShowDialog();

            if (result == true && loginVM.UtilisateurSession != null)
            {
                // Démarrage de la session journalisée
                await Journal.DemarrerSessionAsync(loginVM.UtilisateurSession.Login);

                var mainVM = new MainViewModel
                {
                    UtilisateurConnecte = loginVM.UtilisateurSession
                };

                var mainWin = new MainWindow { DataContext = mainVM };

                // Fin de session à la fermeture
                mainWin.Closed += async (_, _) => await Journal.TerminerSessionAsync();

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
