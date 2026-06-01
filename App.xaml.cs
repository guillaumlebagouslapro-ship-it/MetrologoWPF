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

            // ===== PROFILING DÉMARRAGE (À RETIRER UNE FOIS L'OPTIMISATION FAITE) =====
            // Instrumente chaque grande étape pour identifier les goulots d'étranglement.
            // Logs écrits à la fin dans le journal + un fichier startup_profiling.txt à
            // côté de l'exe (facile à lire/copier).
            var swStartup = System.Diagnostics.Stopwatch.StartNew();
            var lapsStartup = new System.Collections.Generic.List<(string Etape, long MsDebut, long MsFin)>();
            long lastLap = 0;
            void StartupLap(string etape)
            {
                long now = swStartup.ElapsedMilliseconds;
                lapsStartup.Add((etape, lastLap, now));
                lastLap = now;
            }
            // ========================================================================

            Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;
            StartupLap("ShutdownMode set");

            // Vérification des prérequis système (Excel, NI-VISA, ni4882.dll) AVANT de
            // toucher au moindre service qui en dépend — sinon on plante silencieusement
            // sur les postes mal configurés. Si quelque chose manque, on prévient
            // l'utilisateur et on lui propose de continuer en mode dégradé ou quitter.
            var prerequisManquants = VerificationPrerequis.VerifierTout();
            StartupLap("VerificationPrerequis.VerifierTout");
            if (prerequisManquants.Count > 0)
            {
                var dlg = new PrerequisManquantsDialog(prerequisManquants);
                bool? choix = dlg.ShowDialog();
                if (choix != true)
                {
                    // L'utilisateur a choisi de quitter — on coupe court.
                    Application.Current.Shutdown();
                    return;
                }
            }

            // Migration silencieuse des fichiers locaux historiques (à plat dans
            // %LocalAppData%\Metrologo\) vers la nouvelle organisation par sous-dossier
            // (Configuration\, Presets\, Catalogues\, Cache\). Idempotent — ne fait rien
            // si la migration a déjà eu lieu. À appeler en TOUT PREMIER pour que les
            // services qui suivent (DatabaseInitializer, etc.) lisent depuis le nouveau
            // emplacement.
            CheminsMetrologo.MigrerAnciensFichiers();
            StartupLap("MigrerAnciensFichiers");

            // Amorce la structure sur le partage serveur (M:\exe_spe\Data_Metrologo)
            // — crée les sous-dossiers + le fichier maître paths.config.json s'ils
            // n'existent pas. Idempotent + best-effort : si M:\ est indispo (laptop en
            // déplacement), retombe sur les chemins locaux par défaut.
            bool serveurDispo = CheminsMetrologo.AssurerStructureServeur();
            StartupLap("AssurerStructureServeur (M:\\)");

            // Charge les overrides de chemins. Si le poste n'a pas encore de config locale,
            // l'app auto-adopte le fichier maître serveur (bootstrap silencieux d'un poste
            // vierge). Tous les services qui suivent lisent ensuite les chemins partagés.
            CheminsMetrologo.ChargerConfigChemins();
            StartupLap("ChargerConfigChemins");

            // Migration des anciens noms de sous-dossier de modules d'incertitude
            // (ex. TachyContact → TachymetreContact). Idempotent — ne fait rien si déjà
            // migré ou si l'ancien dossier n'existe pas.
            Services.Incertitude.ModulesIncertitudeService.MigrerAnciensNomsDossiers();
            StartupLap("MigrerAnciensNomsDossiers");

            // Chemin local de sauvegarde des rapports : automatique, identique sur tous
            // les postes (C:\Users\Public\Documents\Metrologo_Backup par défaut). On crée
            // le dossier au démarrage s'il n'existe pas — aucune action utilisateur requise.
            // Un admin peut toujours surcharger ce chemin via Admin → Chemins de stockage.
            bool dossierOk = CheminsMetrologo.AssurerDossierMesuresLocal();
            bool premierDemarrage = CheminsMetrologo.EstPremierDemarrage();
            StartupLap("AssurerDossierMesuresLocal + EstPremierDemarrage");

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

            if (serveurDispo)
            {
                Journal.Info(CategorieLog.Systeme, "PARTAGE_SERVEUR_OK",
                    $"Partage serveur accessible : {CheminsMetrologo.BaseServeur} "
                  + $"(master : {CheminsMetrologo.MasterPathsUrl}).");
            }
            else
            {
                Journal.Warn(CategorieLog.Systeme, "PARTAGE_SERVEUR_KO",
                    $"Partage serveur « {CheminsMetrologo.BaseServeur} » indisponible — "
                  + "l'app fonctionnera sur les chemins locaux mis en cache lors du dernier démarrage réussi.");
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

            // Journal centralisé sur fichiers JSON-lines dans le dossier Logs partagé
            // (M:\…\Logs) : les logs de tous les postes y sont écrits en append-only.
            Journal.Configurer(new FichierJournalService());
            StartupLap("Journal.Configurer");

            // Ferme les sessions zombies (laissées ouvertes par un crash, taskkill, ou
            // arrêt brutal lors d'un debug). Sans ça, elles s'accumulent et apparaissent
            // perpétuellement « En cours » dans la liste du journal.
            await Journal.NettoyerSessionsZombiesAsync();
            StartupLap("NettoyerSessionsZombiesAsync (I/O M:\\)");

            // Archive le mois précédent s'il n'est pas déjà archivé. Idempotent et
            // multi-postes-safe (verrou applicatif SQL). Fire-and-forget : ne bloque
            // pas le démarrage de l'app, l'archivage se fait en arrière-plan.
            _ = ArchivesLogsService.ArchiverMoisPrecedentSiNecessaireAsync();

            // Chargement du catalogue local des modèles d'appareils enregistrés par les utilisateurs
            await CatalogueAppareilsService.Instance.ChargerAsync();
            StartupLap("CatalogueAppareilsService.ChargerAsync");

            // Chargement des presets de balayage de stabilité (séédés au 1er démarrage).
            await PresetsStabiliteService.Instance.ChargerAsync();
            StartupLap("PresetsStabiliteService.ChargerAsync");

            // METROLOGO_Stab.xltx est livré avec l'app (extrait de Stab1.xls historique avec
            // graphe pro intégré). Le builder n'est utilisé qu'en filet de sécurité si quelqu'un
            // a supprimé le template — il génère alors une version basique sans graphe pro.
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string stabPath = Path.Combine(baseDir, "Templates", "METROLOGO_Stab.xltx");
                if (!File.Exists(stabPath))
                {
                    StabTemplateBuilder.EnsureExists(
                        Path.Combine(baseDir, "Templates", "METROLOGO.xltx"),
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

            // Reprise des transferts FI vers le réseau qui n'avaient pas pu se faire à la fin
            // des sessions précédentes (M:\ down, latence, etc.). Si la liste est vide ou si
            // M:\ est toujours indispo, no-op. Fire-and-forget : ne bloque pas le démarrage.
            _ = TransfertReseauService.TenterTransfertsEnAttenteAsync();

            // Warm-up ClosedXML : ouvre les 2 templates en arrière-plan pour pré-JIT les
            // assemblies (ClosedXML.dll, DocumentFormat.OpenXml.dll) + déclencher le cache
            // disque OS du fichier .xltx. Sans ce warm-up, la 1ère InitialiserRapportAsync
            // d'une mesure prend ~1 s ; après warm-up, ~25 ms (gain mesuré dans le profiler).
            // Fire-and-forget : ne bloque pas le démarrage de l'app.
            _ = Task.Run(() =>
            {
                try
                {
                    string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                    foreach (var nomTpl in new[] { "METROLOGO.xltx", "METROLOGO_Stab.xltx" })
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

            StartupLap("Vérif template Stab + start Excel COM + warm-up ClosedXML (FaF)");

            // ===== DÉMARRAGE =====
            // L'app démarre directement sur l'écran de sélection utilisateur (menu déroulant
            // alimenté par les comptes locaux). Le MainViewModel orchestre ensuite la suite :
            // sélection Baie/Paillasse → Accueil. La session de journal est démarrée par le
            // MainViewModel au moment où l'utilisateur est choisi.
            var mainVM = new MainViewModel();
            StartupLap("new MainViewModel");
            var mainWin = new MainWindow { DataContext = mainVM };
            StartupLap("new MainWindow (XAML parse)");
            // Note : on NE met PAS « mainWin.Closed += async (_, _) => await Journal.TerminerSessionAsync(); »
            // Le pattern async void laissait l'app s'arrêter avant l'écriture du fichier
            // sessions.json → sessions zombies. La fermeture propre est désormais faite
            // SYNCHRONIQUEMENT dans OnExit (cf. ci-dessous), seul moment garanti d'exécution
            // avant que le process se termine.
            Application.Current.MainWindow = mainWin;
            Application.Current.ShutdownMode = ShutdownMode.OnMainWindowClose;
            mainWin.Show();
            StartupLap("mainWin.Show");

            // ===== ÉCRITURE DU PROFILING (À RETIRER UNE FOIS L'OPTIMISATION FAITE) =====
            swStartup.Stop();
            try
            {
                var sb = new System.Text.StringBuilder();
                sb.AppendLine($"Profiling démarrage Metrologo — Total : {swStartup.ElapsedMilliseconds} ms");
                sb.AppendLine($"Date : {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine();
                sb.AppendLine("Temps(ms)  Delta(ms)  Étape");
                sb.AppendLine("---------- ---------- ----------------------------------------------------------");
                foreach (var (etape, msDeb, msFin) in lapsStartup)
                {
                    long delta = msFin - msDeb;
                    sb.AppendLine($"{msFin,8}   {delta,8}   {etape}");
                }
                string contenu = sb.ToString();

                Journal.Info(CategorieLog.Systeme, "STARTUP_PROFILING",
                    $"Démarrage total {swStartup.ElapsedMilliseconds}ms — voir startup_profiling.txt à côté de l'exe.\n"
                    + contenu);

                // Écrit aussi dans un fichier à côté de l'exe (facile à copier-coller)
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string fichierProfiling = System.IO.Path.Combine(baseDir, "startup_profiling.txt");
                System.IO.File.WriteAllText(fichierProfiling, contenu);
            }
            catch (Exception ex)
            {
                Journal.Warn(CategorieLog.Systeme, "STARTUP_PROFILING_KO",
                    $"Écriture profiling démarrage échouée : {ex.Message}");
            }
            // ========================================================================
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // Termine la session journal FI utilisateur en cours (écrit FIN_SESSION dans
            // Journal_<FI>.txt avec récap mesures effectuées/échouées + durée). Idempotent.
            try { Metrologo.Services.Journal.JournalFIService.TerminerSession("Fermeture application"); }
            catch { /* best-effort */ }

            // Termine la session de journal de manière SYNCHRONE pour garantir que le
            // sessions.json est écrit AVANT l'arrêt du process. Sinon, en async void,
            // l'app peut se terminer avant la fin de l'écriture → la session reste
            // marquée « en cours » indéfiniment dans le viewer du journal.
            try
            {
                Journal.TerminerSessionAsync().GetAwaiter().GetResult();
            }
            catch
            {
                // Best-effort : si l'écriture échoue (partage HS, etc.), on continue
                // l'arrêt. Le nettoyage zombies au prochain démarrage rattrapera.
            }

            // Ferme proprement l'instance Excel cachée — sinon Excel.exe reste en tâche
            // de fond et peut bloquer la sauvegarde de fichiers au prochain lancement.
            ExcelInteropHost.Instance.Dispose();
            base.OnExit(e);
        }
    }
}
