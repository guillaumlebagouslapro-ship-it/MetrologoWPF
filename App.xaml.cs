using Metrologo;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Journal;
using Metrologo.ViewModels;
using Metrologo.Views;
using System;
using System.IO;
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

            // Journal centralisé sur SQL Server : les logs de tous les postes (Baie,
            // Paillasse, dev) atterrissent dans la base Metrologo et sont consultables
            // par l'admin depuis n'importe quelle machine.
            Journal.Configurer(new SqlServerJournalService());

            // Ferme les sessions zombies (laissées ouvertes par un crash, taskkill, ou
            // arrêt brutal lors d'un debug). Sans ça, elles s'accumulent et apparaissent
            // perpétuellement « En cours » dans la liste du journal.
            await Journal.NettoyerSessionsZombiesAsync();

            // Archive le mois précédent s'il n'est pas déjà archivé. Idempotent et
            // multi-postes-safe (verrou applicatif SQL). Fire-and-forget : ne bloque
            // pas le démarrage de l'app, l'archivage se fait en arrière-plan.
            _ = ArchivesLogsService.ArchiverMoisPrecedentSiNecessaireAsync();

            // Chargement du catalogue local des modèles d'appareils enregistrés par les utilisateurs
            await CatalogueAppareilsService.Instance.ChargerAsync();

            // Chargement des presets de balayage de stabilité (séédés au 1er démarrage).
            await PresetsStabiliteService.Instance.ChargerAsync();

            // METROLOGO_Stab.xltm est livré avec l'app (extrait de Stab1.xls historique avec
            // graphe pro intégré). Le builder n'est utilisé qu'en filet de sécurité si quelqu'un
            // a supprimé le template — il génère alors une version basique sans graphe pro.
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string stabPath = Path.Combine(baseDir, "Templates", "METROLOGO_Stab.xltm");
                if (!File.Exists(stabPath))
                {
                    StabTemplateBuilder.EnsureExists(
                        Path.Combine(baseDir, "Templates", "METROLOGO.xltm"),
                        stabPath);
                }
            }
            catch (Exception ex)
            {
                Journal.Warn(CategorieLog.Excel, "TEMPLATE_STAB_BUILD",
                    $"Construction du template Stabilité échouée : {ex.Message}.");
            }

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
