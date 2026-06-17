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

            // ===== PROFILING DÉMARRAGE (À RETIRER APRÈS OPTIMISATION) =====
            // Chronomètre chaque étape ; résultat dans le journal + startup_profiling.txt.
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

            // Attribue automatiquement la fenêtre active comme Owner de toute modale sans
            // Owner explicite. Sans ça, ShowDialog peut passer derrière la principale
            // (désactivée → app gelée en apparence). Windows maintient alors la modale
            // au-dessus et la fait clignoter si on clique sur la principale.
            EventManager.RegisterClassHandler(typeof(Window), Window.LoadedEvent,
                new RoutedEventHandler(AttribuerOwnerAutomatique));

            // Adapte chaque fenêtre à la zone de travail : rétrécit si besoin et la
            // recale pour qu'elle soit entièrement visible. Le ScrollViewer prend le
            // relais pour le contenu qui ne tient plus.
            EventManager.RegisterClassHandler(typeof(Window), Window.LoadedEvent,
                new RoutedEventHandler(AdapterTailleAEcran));

            // Prérequis (Excel, NI-VISA, ni4882.dll) vérifiés avant les services qui en
            // dépendent — sinon échec silencieux. Propose mode dégradé ou arrêt.
            var prerequisManquants = VerificationPrerequis.VerifierTout();
            StartupLap("VerificationPrerequis.VerifierTout");
            if (prerequisManquants.Count > 0)
            {
                var dlg = new PrerequisManquantsDialog(prerequisManquants);
                bool? choix = dlg.ShowDialog();
                if (choix != true)
                {
                    // l'utilisateur a choisi de quitter
                    Application.Current.Shutdown();
                    return;
                }
            }

            // Migration des anciens fichiers plats (%LocalAppData%\Metrologo\) vers les
            // sous-dossiers (Configuration\, Presets\, Catalogues\, Cache\). Idempotent.
            // À appeler avant tous les services qui lisent ces chemins.
            CheminsMetrologo.MigrerAnciensFichiers();
            StartupLap("MigrerAnciensFichiers");

            // Crée la structure sur M:\exe_spe\Data_Metrologo (sous-dossiers +
            // paths.config.json) si absente. Idempotent ; retombe sur les chemins
            // locaux si M:\ est indisponible.
            bool serveurDispo = CheminsMetrologo.AssurerStructureServeur();
            StartupLap("AssurerStructureServeur (M:\\)");

            // Partage inaccessible → pop-up bloquante : Rafraîchir relance le test,
            // Fermer arrête l'app. Le dialogue ne se clôt qu'à la reconnexion réussie.
            if (!serveurDispo)
            {
                var dlgReseau = new ReseauIndisponibleDialog(
                    CheminsMetrologo.BaseServeur,
                    CheminsMetrologo.AssurerStructureServeur);
                bool? choixReseau = dlgReseau.ShowDialog();
                if (choixReseau != true)
                {
                    // L'utilisateur a choisi de fermer l'application.
                    Application.Current.Shutdown();
                    return;
                }
                serveurDispo = true;
                StartupLap("ReseauIndisponibleDialog (rafraîchi)");
            }

            // Charge les overrides de chemins. Poste vierge : adopte le maître serveur
            // silencieusement (bootstrap). Tous les services suivants lisent les chemins partagés.
            CheminsMetrologo.ChargerConfigChemins();
            StartupLap("ChargerConfigChemins");

            // Renommage des anciens dossiers de modules d'incertitude (ex. TachyContact
            // → TachymetreContact). Idempotent.
            Services.Incertitude.ModulesIncertitudeService.MigrerAnciensNomsDossiers();
            StartupLap("MigrerAnciensNomsDossiers");

            // Dossier local de sauvegarde des rapports (C:\Users\Public\Documents\Metrologo_Backup).
            // Créé au démarrage si absent. Overridable via Admin → Chemins de stockage.
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

            // Premier démarrage : informe l'utilisateur du dossier créé et ajoute un
            // raccourci Bureau. Affiché une seule fois (flag dans Configuration\).
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

            // Journal JSON-lines centralisé sur M:\…\Logs (append-only, tous les postes).
            Journal.Configurer(new FichierJournalService());
            StartupLap("Journal.Configurer");

            // (Diagnostic SQL retiré : suivi Besançon 100 % fichier txt. Évite le timeout
            // de 6 s au démarrage quand SVR-OR est injoignable.)

            // Sessions zombies (crash, taskkill, debug brutal) : sans ce nettoyage elles
            // s'accumulent et restent « En cours » dans le journal indéfiniment.
            await Journal.NettoyerSessionsZombiesAsync();
            StartupLap("NettoyerSessionsZombiesAsync (I/O M:\\)");

            // Archive le mois précédent si besoin. Idempotent, multi-postes-safe.
            // Fire-and-forget.
            _ = ArchivesLogsService.ArchiverMoisPrecedentSiNecessaireAsync();

            // Catalogue local des modèles d'appareils.
            await CatalogueAppareilsService.Instance.ChargerAsync();
            StartupLap("CatalogueAppareilsService.ChargerAsync");

            // Fréquencemètres legacy (EIP / Racal / Stanford, sans *IDN?) : seed en mémoire
            // pour la sélection en « adresses fixes ». Idempotent.
            Metrologo.Services.Catalogue.SeedLegacyAppareils.EnsureSeeded();
            StartupLap("SeedLegacyAppareils.EnsureSeeded");

            // Presets de balayage de stabilité (séédés au 1er démarrage).
            await PresetsStabiliteService.Instance.ChargerAsync();
            StartupLap("PresetsStabiliteService.ChargerAsync");

            // METROLOGO_Stab.xlsx livré avec l'app (issu de Stab1.xls, graphe pro intégré).
            // Builder déclenché uniquement si le template a été supprimé → version basique sans graphe.
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string stabPath = Path.Combine(baseDir, "Templates", "METROLOGO_Stab.xlsx");
                if (!File.Exists(stabPath))
                {
                    StabTemplateBuilder.EnsureExists(
                        Path.Combine(baseDir, "Templates", "METROLOGO.xlsx"),
                        stabPath);
                }
            }
            catch (Exception ex)
            {
                Journal.Warn(CategorieLog.Excel, "TEMPLATE_STAB_BUILD",
                    $"Construction du template Stabilité échouée : {ex.Message}.");
            }

            // Instance Excel cachée pré-démarrée (héritage Delphi) : plus de coût COM
            // dans la boucle chaude au lancement d'une mesure.
            _ = ExcelInteropHost.Instance.DemarrerAsync();

            // Reprend les transferts FI en attente (M:\ down ou latence lors des sessions
            // précédentes). No-op si liste vide ou M:\ toujours indispo. Fire-and-forget.
            _ = TransfertReseauService.TenterTransfertsEnAttenteAsync();

            // Tâche quotidienne Besançon : FTP → valeurs_besancon.txt. Le marqueur partagé
            // empêche les doublons inter-postes. Fire-and-forget.
            try { Metrologo.Services.Besancon.BesanconScheduler.Demarrer(); }
            catch (Exception exBes)
            {
                Journal.Warn(CategorieLog.Systeme, "BESANCON_DEMARRAGE_KO",
                    $"Démarrage de la tâche Besançon échoué : {exBes.Message}");
            }

            // (Rattrapage hebdo SQL retiré : moyennes Besançon recalculées à la volée depuis
            // le txt. Plus de timeout ni SqlException quand SVR-OR est injoignable.)

            // Warm-up ClosedXML : ouvre les 2 templates en arrière-plan pour pré-JIT
            // (ClosedXML.dll, OpenXml.dll) et amorcer le cache disque. Sans ça, la 1ère
            // InitialiserRapportAsync prend ~1 s ; après warm-up ~25 ms. Fire-and-forget.
            _ = Task.Run(() =>
            {
                try
                {
                    string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                    foreach (var nomTpl in new[] { "METROLOGO.xlsx", "METROLOGO_Stab.xlsx" })
                    {
                        string chemin = Path.Combine(baseDir, "Templates", nomTpl);
                        if (!File.Exists(chemin)) continue;
                        using var wb = new ClosedXML.Excel.XLWorkbook(chemin);
                        // Force le lazy-load complet (ClosedXML ne charge qu'au premier accès).
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
            // L'app s'ouvre sur l'écran de sélection utilisateur. MainViewModel orchestre
            // la suite (Baie/Paillasse → Accueil) et démarre la session journal au choix.
            var mainVM = new MainViewModel();
            StartupLap("new MainViewModel");
            var mainWin = new MainWindow { DataContext = mainVM };
            StartupLap("new MainWindow (XAML parse)");
            // Pas de mainWin.Closed += async void : l'app s'arrêtait avant l'écriture de
            // sessions.json → zombies. Fermeture faite de façon synchrone dans OnExit.
            Application.Current.MainWindow = mainWin;
            Application.Current.ShutdownMode = ShutdownMode.OnMainWindowClose;
            mainWin.Show();
            StartupLap("mainWin.Show");

            // Surveille les changements admin inter-postes → toast non-bloquant (rubidium,
            // chemins, modules, catalogue, utilisateurs). Sur le thread UI.
            try { Metrologo.Services.Journal.NotificationsAdminWatcher.Demarrer(); }
            catch { /* la notif inter-postes n'est pas critique */ }

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

                // Copie dans un fichier à côté de l'exe (facile à partager).
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

        /// <summary>
        /// Handler de classe sur Loaded : attribue la fenêtre active comme Owner à toute
        /// Window sans Owner explicite. Gère les dialogues imbriqués (l'actif est le
        /// parent, pas la MainWindow). Best-effort : échec = comportement habituel.
        /// </summary>
        private static void AttribuerOwnerAutomatique(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is not Window fenetre) return;

                // Thread UI secondaire (ex. StopMesureFloatingWindow) : ne rien faire.
                // Current.MainWindow et Current.Windows appartiennent au thread principal ;
                // y accéder cross-thread lève InvalidOperationException (plantage STOP).
                var app = Application.Current;
                if (app == null || app.Dispatcher != fenetre.Dispatcher) return;

                if (fenetre == app.MainWindow) return;
                if (fenetre.Owner != null) return;
                // Topmost (toasts) : pas d'Owner, sinon la fenêtre parente les fermerait.
                if (fenetre.Topmost) return;
                Window? proprietaire = null;
                foreach (Window w in app.Windows)
                {
                    if (w != fenetre && w.IsActive && w.IsVisible) { proprietaire = w; break; }
                }
                proprietaire ??= app.MainWindow;

                if (proprietaire == null || proprietaire == fenetre
                    || !proprietaire.IsVisible || proprietaire.Owner == fenetre)
                    return;

                try
                {
                    fenetre.Owner = proprietaire;
                }
                catch (InvalidOperationException)
                {
                    // ShowDialog déjà appelé : WPF refuse Owner à ce stade. On le pose
                    // via GWLP_HWNDPARENT (Win32) — même effet z-order.
                    nint hwnd = new System.Windows.Interop.WindowInteropHelper(fenetre).Handle;
                    nint hOwner = new System.Windows.Interop.WindowInteropHelper(proprietaire).Handle;
                    if (hwnd != 0 && hOwner != 0)
                        SetWindowLongPtr(hwnd, GWLP_HWNDPARENT, hOwner);
                }
            }
            catch
            {
                // Best-effort : sans Owner la fenêtre s'ouvre quand même, comme avant.
            }
        }

        /// <summary>
        /// Handler de classe sur Loaded : adapte la fenêtre à la zone de travail.
        /// Grand écran : agrandissement proportionnel (plafonné à ×1.4).
        /// Petit écran : réduction pour tenir, puis recalage entièrement visible.
        /// Ignoré pour les fenêtres maximisées et Topmost.
        /// </summary>
        private static void AdapterTailleAEcran(object sender, RoutedEventArgs e)
        {
            if (sender is not Window fenetre) return;
            if (fenetre.WindowState != WindowState.Normal) return;
            if (fenetre.Topmost) return;

            try
            {
                Rect zone = SystemParameters.WorkArea;
                const double marge = 16;
                double maxLargeur = zone.Width - marge;
                double maxHauteur = zone.Height - marge;

                double largeurInitiale = double.IsNaN(fenetre.Width) ? fenetre.ActualWidth : fenetre.Width;
                double hauteurInitiale = double.IsNaN(fenetre.Height) ? fenetre.ActualHeight : fenetre.Height;

                // Référence de conception : 1920×1040 WPF (Full HD). Facteur toujours >= 1 ;
                // la réduction passe par maxLargeur/maxHauteur, pas par ce facteur.
                const double refLargeur = 1920, refHauteur = 1040;
                double facteur = Math.Min(zone.Width / refLargeur, zone.Height / refHauteur);
                facteur = Math.Clamp(facteur, 1.0, 1.4);

                double largeur = Math.Min(largeurInitiale * facteur, maxLargeur);
                double hauteur = Math.Min(hauteurInitiale * facteur, maxHauteur);

                bool retaillee = Math.Abs(largeur - largeurInitiale) > 0.5
                              || Math.Abs(hauteur - hauteurInitiale) > 0.5;
                if (retaillee)
                {
                    // Conserve le centre d'origine pour que le redimensionnement ne décale pas.
                    double centreX = fenetre.Left + largeurInitiale / 2;
                    double centreY = fenetre.Top + hauteurInitiale / 2;
                    fenetre.Width = largeur;
                    fenetre.Height = hauteur;
                    fenetre.Left = centreX - largeur / 2;
                    fenetre.Top = centreY - hauteur / 2;
                }

                // Recale dans la zone de travail (après retaille ou ouverture à cheval).
                if (retaillee || fenetre.Left < zone.Left || fenetre.Top < zone.Top
                    || fenetre.Left + largeur > zone.Right || fenetre.Top + hauteur > zone.Bottom)
                {
                    fenetre.Left = Math.Max(zone.Left, Math.Min(fenetre.Left, zone.Right - largeur));
                    fenetre.Top = Math.Max(zone.Top, Math.Min(fenetre.Top, zone.Bottom - hauteur));
                }
            }
            catch
            {
                // Best-effort : en cas de souci la fenêtre garde sa taille déclarée.
            }
        }

        private const int GWLP_HWNDPARENT = -8;

        private static nint SetWindowLongPtr(nint hWnd, int nIndex, nint dwNewLong) =>
            nint.Size == 8
                ? SetWindowLongPtr64(hWnd, nIndex, dwNewLong)
                : SetWindowLong32(hWnd, nIndex, (int)dwNewLong);

        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
        private static extern nint SetWindowLongPtr64(nint hWnd, int nIndex, nint dwNewLong);

        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SetWindowLongW")]
        private static extern int SetWindowLong32(nint hWnd, int nIndex, int dwNewLong);

        protected override void OnExit(ExitEventArgs e)
        {
            // FIN_SESSION dans Journal_<FI>.txt (récap mesures + durée). Idempotent.
            try { Metrologo.Services.Journal.JournalFIService.TerminerSession("Fermeture application"); }
            catch { /* best-effort */ }

            // Fermeture synchrone obligatoire : sessions.json doit être écrit avant
            // l'arrêt du process. En async void l'app pouvait se terminer avant
            // l'écriture → session restait « en cours » indéfiniment.
            try
            {
                Journal.TerminerSessionAsync().GetAwaiter().GetResult();
            }
            catch
            {
                // Best-effort : l'arrêt continue ; les zombies seront nettoyés au démarrage suivant.
            }

            // Ferme l'instance Excel cachée — sinon Excel.exe reste en tâche de fond.
            ExcelInteropHost.Instance.Dispose();
            base.OnExit(e);
        }
    }
}
