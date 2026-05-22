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

            // Migration silencieuse des fichiers locaux historiques (à plat dans
            // %LocalAppData%\Metrologo\) vers la nouvelle organisation par sous-dossier
            // (Configuration\, Presets\, Catalogues\, Cache\). Idempotent — ne fait rien
            // si la migration a déjà eu lieu. À appeler en TOUT PREMIER pour que les
            // services qui suivent (DatabaseInitializer, etc.) lisent depuis le nouveau
            // emplacement.
            CheminsMetrologo.MigrerAnciensFichiers();

            // Charge les overrides de chemins (Configuration\paths.config.json) — permet à
            // l'admin de pointer Incertitudes / Presets / Catalogues vers un partage réseau
            // commun. Sans fichier, les services utilisent les chemins locaux par défaut.
            CheminsMetrologo.ChargerConfigChemins();

            // Migration des anciens noms de sous-dossier de modules d'incertitude
            // (ex. TachyContact → TachymetreContact). Idempotent — ne fait rien si déjà
            // migré ou si l'ancien dossier n'existe pas.
            Services.Incertitude.ModulesIncertitudeService.MigrerAnciensNomsDossiers();

            // Chemin local de sauvegarde des rapports : automatique, identique sur tous
            // les postes (C:\Users\Public\Documents\Metrologo_Backup par défaut). On crée
            // le dossier au démarrage s'il n'existe pas — aucune action utilisateur requise.
            // Un admin peut toujours surcharger ce chemin via Admin → Chemins de stockage.
            bool dossierOk = CheminsMetrologo.AssurerDossierMesuresLocal();
            bool premierDemarrage = CheminsMetrologo.EstPremierDemarrage();

            if (dossierOk)
            {
                Journal.Info(CategorieLog.Systeme, "DOSSIER_MESURES_LOCAL_PRET",
                    $"Dossier local de sauvegarde prêt : {CheminsMetrologo.MesuresLocal}");
            }
            else
            {
                Journal.Warn(CategorieLog.Systeme, "DOSSIER_MESURES_LOCAL_KO",
                    $"Impossible de créer le dossier local de sauvegarde « {CheminsMetrologo.MesuresLocal} ». "
                  + "La duplication des rapports échouera silencieusement — un admin doit corriger via "
                  + "Admin → Chemins de stockage.");
            }

            // Message d'accueil au tout premier démarrage : informe l'utilisateur de
            // l'emplacement créé + ajoute un raccourci sur son Bureau pour y accéder en
            // un clic. Affiché une seule fois (flag persistant dans Configuration\).
            if (premierDemarrage && dossierOk)
            {
                string? raccourci = CheminsMetrologo.CreerRaccourciBureauMesuresLocal();
                string msg =
                    "Bienvenue sur Metrologo.\n\n"
                  + "Un dossier local de sauvegarde des rapports a été créé automatiquement :\n\n"
                  + $"   {CheminsMetrologo.MesuresLocal}\n\n"
                  + "Chaque mesure y sera dupliquée automatiquement après son rapport principal — "
                  + "tu retrouveras donc tous tes Excel à cet endroit même en cas de coupure réseau.";
                if (!string.IsNullOrEmpty(raccourci))
                {
                    msg += "\n\nUn raccourci « Mesures Metrologo (local) » a également été ajouté "
                         + "sur ton Bureau pour y accéder rapidement.";
                }

                Journal.Info(CategorieLog.Systeme, "PREMIER_DEMARRAGE",
                    $"Premier démarrage Metrologo sur ce poste — dossier local : {CheminsMetrologo.MesuresLocal} "
                  + (raccourci != null ? $"; raccourci : {raccourci}" : "; raccourci : non créé"));

                MessageBox.Show(msg, "Metrologo — premier démarrage",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                CheminsMetrologo.MarquerPremierDemarrageEffectue();
            }

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

            // Warm-up ClosedXML : ouvre les 2 templates en arrière-plan pour pré-JIT les
            // assemblies (ClosedXML.dll, DocumentFormat.OpenXml.dll) + déclencher le cache
            // disque OS du fichier .xltm. Sans ce warm-up, la 1ère InitialiserRapportAsync
            // d'une mesure prend ~1 s ; après warm-up, ~25 ms (gain mesuré dans le profiler).
            // Fire-and-forget : ne bloque pas le démarrage de l'app.
            _ = Task.Run(() =>
            {
                try
                {
                    string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                    foreach (var nomTpl in new[] { "METROLOGO.xltm", "METROLOGO_Stab.xltm" })
                    {
                        string chemin = Path.Combine(baseDir, "Templates", nomTpl);
                        if (!File.Exists(chemin)) continue;
                        using var wb = new ClosedXML.Excel.XLWorkbook(chemin);
                        // Touche chaque feuille pour forcer le lazy-load complet (sinon
                        // ClosedXML ne charge réellement la feuille qu'au premier accès).
                        foreach (var ws in wb.Worksheets)
                        {
                            _ = ws.Name;
                            _ = ws.RowsUsed().Count();
                        }
                    }
                    Journal.Info(CategorieLog.Excel, "WARMUP_OK",
                        "Warm-up ClosedXML terminé — 1ère mesure aura un démarrage instantané.");
                }
                catch (Exception ex)
                {
                    Journal.Warn(CategorieLog.Excel, "WARMUP_KO",
                        $"Warm-up ClosedXML échoué : {ex.Message}");
                }
            });

            // ===== MODE DEV (publication) =====
            // L'app saute l'écran de login : auto-connexion en tant qu'utilisateur "dev"
            // avec le rôle Administrateur. L'utilisateur arrive directement sur l'écran
            // de sélection Baie/Paillasse, qu'il garde (= choix conscient du poste).
            //
            // Pour réactiver l'écran de login en mode production : remettre l'ancien flux
            // (LoginWindow + AuthService) à la place de ce bloc — ou conditionner par un
            // flag #DEFINE / un paramètre de release au moment du build.
            var fakeUser = new Utilisateur
            {
                Login = "dev",
                Nom = "Mode Dev",
                Prenom = "Test",
                Role = RoleUtilisateur.Administrateur,
                Actif = true,
                DateCreation = DateTime.UtcNow
            };

            Journal.Warn(CategorieLog.Systeme, "DEV_BYPASS",
                "Mode dev : login bypassé, auto-connexion utilisateur=dev (admin). "
              + "Sélection Baie/Paillasse conservée pour choix conscient.");

            await Journal.DemarrerSessionAsync(fakeUser.Login);

            var mainVMDev = new MainViewModel { UtilisateurConnecte = fakeUser };
            // PAS de BypassSelectionPoste → l'écran de sélection s'affiche normalement.

            var mainWinDev = new MainWindow { DataContext = mainVMDev };
            mainWinDev.Closed += async (_, _) => await Journal.TerminerSessionAsync();
            Application.Current.MainWindow = mainWinDev;
            Application.Current.ShutdownMode = ShutdownMode.OnMainWindowClose;
            mainWinDev.Show();
            return;

            // ===== FLUX LOGIN PRODUCTION (désactivé en mode dev) =====
            // Code conservé en mort pour réactivation rapide en passage prod.
#pragma warning disable CS0162   // Unreachable code
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
