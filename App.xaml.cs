using Metrologo;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Config;
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

            // Chargement de la configuration des appareils IEEE depuis Metrologo.ini
            ChargerConfigAppareils();

            // Chargement du catalogue local des modèles d'appareils enregistrés par les utilisateurs
            await CatalogueAppareilsService.Instance.ChargerAsync();

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

        private static void ChargerConfigAppareils()
        {
            var chemin = Preferences.CheminFichierIni;
            try
            {
                var config = ConfigAppareilsLoader.Charger(chemin);
                EtatApplication.ConfigAppareils = config;

                Journal.Info(CategorieLog.Configuration, "ChargementIni",
                    $"Configuration des appareils chargée depuis « {chemin} ».",
                    new
                    {
                        Stanford = $"@{config.Stanford.Adresse} ({config.Stanford.Gates.Count} gates)",
                        Racal    = $"@{config.Racal.Adresse} ({config.Racal.Gates.Count} gates)",
                        EIP      = $"@{config.Eip.Adresse} ({config.Eip.Gates.Count} gates)",
                        Mux      = config.Mux == null ? "absent" : $"@{config.Mux.Adresse}"
                    });

                foreach (var avert in config.Avertissements)
                    Journal.Warn(CategorieLog.Configuration, "ChargementIni", avert);
            }
            catch (Exception ex)
            {
                Journal.Erreur(CategorieLog.Configuration, "ChargementIni",
                    $"Échec du chargement de « {chemin} » : {ex.Message}");
                // Non bloquant pour le moment (driver IEEE pas encore branché).
                // À durcir quand le VISA réel sera en place.
            }
        }
    }
}
