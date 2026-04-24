using Metrologo;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Journal;
using Metrologo.ViewModels;
using Metrologo.Views;
using System;
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

            // Chargement du catalogue local des modèles d'appareils enregistrés par les utilisateurs
            await CatalogueAppareilsService.Instance.ChargerAsync();

            // Démarre une instance Excel cachée en arrière-plan dès le lancement de l'app
            // (comportement hérité du Delphi) pour que l'ouverture du classeur au démarrage
            // d'une mesure soit instantanée — plus de temps COM à payer dans la boucle chaude.
            _ = ExcelInteropHost.Instance.DemarrerAsync();

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

        protected override void OnExit(ExitEventArgs e)
        {
            // Ferme proprement l'instance Excel cachée — sinon Excel.exe reste en tâche de fond
            // et peut bloquer la sauvegarde de fichiers au prochain lancement.
            ExcelInteropHost.Instance.Dispose();
            base.OnExit(e);
        }
    }
}
