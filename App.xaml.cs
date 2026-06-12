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
            // Chronomètre chaque grande étape pour trouver les goulots. Résultat écrit à
            // la fin dans le journal + startup_profiling.txt à côté de l'exe.
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

            // Toute fenêtre secondaire ouverte sans Owner explicite reçoit automatiquement
            // la fenêtre active comme propriétaire. Sans Owner, une fenêtre modale
            // (ShowDialog) peut passer DERRIÈRE la fenêtre principale (clic, Alt-Tab…) :
            // la principale est désactivée par la modale devenue invisible → l'app semble
            // gelée. Avec le lien de propriété, Windows maintient la modale au-dessus de
            // son propriétaire et la ramène/clignote quand on clique sur la principale.
            EventManager.RegisterClassHandler(typeof(Window), Window.LoadedEvent,
                new RoutedEventHandler(AttribuerOwnerAutomatique));

            // Les fenêtres ont des tailles fixes pensées pour un grand écran (ex.
            // Configuration : 860 px de haut) — sur un petit écran elles débordent du
            // cadre. À l'ouverture de chaque fenêtre, on la rétrécit si besoin pour
            // qu'elle tienne dans la zone de travail (écran moins barre des tâches)
            // et on la repositionne entièrement visible. Le contenu scrollable
            // (ScrollViewer) prend le relais pour la partie qui ne tient plus.
            EventManager.RegisterClassHandler(typeof(Window), Window.LoadedEvent,
                new RoutedEventHandler(AdapterTailleAEcran));

            // Vérif des prérequis (Excel, NI-VISA, ni4882.dll) AVANT de toucher aux services
            // qui en dépendent, sinon ça plante silencieusement sur les postes mal configurés.
            // S'il manque quelque chose, on propose de continuer en mode dégradé ou de quitter.
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

            // Migration des anciens fichiers locaux (à plat dans %LocalAppData%\Metrologo\)
            // vers les sous-dossiers Configuration\, Presets\, Catalogues\, Cache\. Idempotent.
            // À appeler en TOUT PREMIER pour que les services suivants lisent au bon endroit.
            CheminsMetrologo.MigrerAnciensFichiers();
            StartupLap("MigrerAnciensFichiers");

            // Amorce la structure sur le partage serveur (M:\exe_spe\Data_Metrologo)
            // — crée les sous-dossiers + le fichier maître paths.config.json s'ils
            // n'existent pas. Idempotent + best-effort : si M:\ est indispo (laptop en
            // déplacement), retombe sur les chemins locaux par défaut.
            bool serveurDispo = CheminsMetrologo.AssurerStructureServeur();
            StartupLap("AssurerStructureServeur (M:\\)");

            // Partage réseau inaccessible → pop-up bloquante : l'utilisateur peut
            // « Rafraîchir » (relance le test d'accès, ex. après avoir rebranché le
            // câble / reconnecté le VPN) ou fermer l'application. Le dialogue ne se
            // ferme en succès que lorsque le partage redevient joignable.
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

            // (Diagnostic SQL Metrologo retiré : le suivi Besançon est désormais 100 % fichier txt,
            // plus aucune base SQL. Évitait sinon un timeout de connexion de 6 s au démarrage —
            // visible en debug sous forme de SqlException — quand SVR-OR est injoignable.)

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

            // Seed des fréquencemètres legacy (EIP / Racal / Stanford) qui ne répondent pas à
            // *IDN? : injectés en mémoire (idempotent) pour la sélection en « adresses fixes ».
            Metrologo.Services.Catalogue.SeedLegacyAppareils.EnsureSeeded();
            StartupLap("SeedLegacyAppareils.EnsureSeeded");

            // Chargement des presets de balayage de stabilité (séédés au 1er démarrage).
            await PresetsStabiliteService.Instance.ChargerAsync();
            StartupLap("PresetsStabiliteService.ChargerAsync");

            // METROLOGO_Stab.xlsx est livré avec l'app (extrait de Stab1.xls historique avec
            // graphe pro intégré). Le builder n'est utilisé qu'en filet de sécurité si quelqu'un
            // a supprimé le template — il génère alors une version basique sans graphe pro.
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

            // Démarre une instance Excel cachée en arrière-plan dès le lancement de l'app
            // (comportement hérité du Delphi) pour que l'ouverture du classeur au démarrage
            // d'une mesure soit instantanée — plus de temps COM à payer dans la boucle chaude.
            _ = ExcelInteropHost.Instance.DemarrerAsync();

            // Reprise des transferts FI vers le réseau qui n'avaient pas pu se faire à la fin
            // des sessions précédentes (M:\ down, latence, etc.). Si la liste est vide ou si
            // M:\ est toujours indispo, no-op. Fire-and-forget : ne bloque pas le démarrage.
            _ = TransfertReseauService.TenterTransfertsEnAttenteAsync();

            // Tâche quotidienne Besançon : récupération FTP du fichier corrigé → fichier txt
            // cumulatif (valeurs_besancon.txt). Active sur tous les postes par défaut ; un marqueur
            // partagé évite que plusieurs postes téléchargent le même jour. Fire-and-forget.
            try { Metrologo.Services.Besancon.BesanconScheduler.Demarrer(); }
            catch (Exception exBes)
            {
                Journal.Warn(CategorieLog.Systeme, "BESANCON_DEMARRAGE_KO",
                    $"Démarrage de la tâche Besançon échoué : {exBes.Message}");
            }

            // (Rattrapage hebdo SQL retiré : les moyennes hebdomadaires Besançon sont désormais
            // recalculées à la volée depuis le fichier txt — plus aucun accès SQL au démarrage,
            // donc plus de timeout de 6 s ni de SqlException quand SVR-OR est injoignable.)

            // Warm-up ClosedXML : ouvre les 2 templates en arrière-plan pour pré-JIT les
            // assemblies (ClosedXML.dll, DocumentFormat.OpenXml.dll) + déclencher le cache
            // disque OS du fichier .xlsx. Sans ce warm-up, la 1ère InitialiserRapportAsync
            // d'une mesure prend ~1 s ; après warm-up, ~25 ms (gain mesuré dans le profiler).
            // Fire-and-forget : ne bloque pas le démarrage de l'app.
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

            // Surveillance des changements admin faits depuis un AUTRE poste → toast non-bloquant
            // (rubidium, chemins, modules d'incertitude, catalogue, utilisateurs…). Thread UI.
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

        /// <summary>
        /// Handler de classe appelé au Loaded de CHAQUE Window de l'app. Si la fenêtre
        /// n'a pas d'Owner (cas de tous les ShowDialog historiques du code), on lui
        /// attribue la fenêtre actuellement active — typiquement celle depuis laquelle
        /// l'utilisateur vient d'ouvrir le dialogue (gère aussi les dialogues imbriqués :
        /// l'actif est alors le dialogue parent, pas la MainWindow). Best-effort : si
        /// l'attribution échoue (fenêtre propriétaire en cours de fermeture, cycle…),
        /// on retombe sur le comportement d'avant, sans casser l'ouverture.
        /// </summary>
        private static void AttribuerOwnerAutomatique(object sender, RoutedEventArgs e)
        {
            if (sender is not Window fenetre) return;
            if (fenetre == Application.Current.MainWindow) return;
            if (fenetre.Owner != null) return;
            // Les fenêtres Topmost (toasts, bouton stop flottant, saisies post-mesure) ne
            // peuvent pas se perdre derrière — et un Owner les fermerait avec la fenêtre
            // qui les a ouvertes, ce qu'on ne veut pas pour un toast.
            if (fenetre.Topmost) return;

            try
            {
                Window? proprietaire = null;
                foreach (Window w in Application.Current.Windows)
                {
                    if (w != fenetre && w.IsActive && w.IsVisible) { proprietaire = w; break; }
                }
                proprietaire ??= Application.Current.MainWindow;

                if (proprietaire == null || proprietaire == fenetre
                    || !proprietaire.IsVisible || proprietaire.Owner == fenetre)
                    return;

                try
                {
                    fenetre.Owner = proprietaire;
                }
                catch (InvalidOperationException)
                {
                    // Fenêtre déjà affichée via ShowDialog : WPF interdit de définir Owner
                    // à ce stade. On pose alors le propriétaire directement au niveau Win32
                    // (GWLP_HWNDPARENT) — même effet sur le z-order : la modale reste
                    // au-dessus de son propriétaire et clignote si on clique dessous.
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
        /// Handler de classe appelé au Loaded de CHAQUE Window : adapte la fenêtre à la
        /// zone de travail de l'écran, dans les deux sens.
        ///   • Grand écran : la fenêtre est agrandie proportionnellement (les tailles
        ///     déclarées dans les XAML sont conçues pour un écran Full HD), plafonné à
        ///     ×1.4 pour ne pas diluer le contenu sur un très grand écran.
        ///   • Petit écran : la fenêtre est rétrécie pour tenir dans la zone de travail
        ///     (le contenu scrollable prend le relais), puis recalée entièrement visible.
        /// Ne touche pas aux fenêtres maximisées ni aux fenêtres auto-positionnées
        /// (toasts Topmost qui gèrent leur propre placement).
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

                // Facteur d'agrandissement proportionnel à l'écran. Référence de
                // conception : zone de travail Full HD (1920×1040 unités WPF). Jamais
                // < 1 ici — la réduction passe par le plafonnement maxLargeur/maxHauteur
                // ci-dessous, qui ne rétrécit que ce qui déborde réellement.
                const double refLargeur = 1920, refHauteur = 1040;
                double facteur = Math.Min(zone.Width / refLargeur, zone.Height / refHauteur);
                facteur = Math.Clamp(facteur, 1.0, 1.4);

                double largeur = Math.Min(largeurInitiale * facteur, maxLargeur);
                double hauteur = Math.Min(hauteurInitiale * facteur, maxHauteur);

                bool retaillee = Math.Abs(largeur - largeurInitiale) > 0.5
                              || Math.Abs(hauteur - hauteurInitiale) > 0.5;
                if (retaillee)
                {
                    // Conserve le centre d'origine (les fenêtres s'ouvrent centrées) pour
                    // que l'agrandissement/réduction ne décale pas la fenêtre.
                    double centreX = fenetre.Left + largeurInitiale / 2;
                    double centreY = fenetre.Top + hauteurInitiale / 2;
                    fenetre.Width = largeur;
                    fenetre.Height = hauteur;
                    fenetre.Left = centreX - largeur / 2;
                    fenetre.Top = centreY - hauteur / 2;
                }

                // Recale la fenêtre pour qu'elle soit entièrement dans la zone de travail
                // (après retaille, ou si elle a été ouverte à cheval sur un bord).
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
