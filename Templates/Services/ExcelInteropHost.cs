using Metrologo.Models;
using Metrologo.Services.Incertitude;
using Metrologo.Services.Journal;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services
{
    /// <summary>
    /// Hôte Excel partagé par toute l'application : une seule instance <c>Excel.Application</c>
    /// est lancée en arrière-plan au démarrage (invisible). Chaque mesure réutilise cette instance
    /// pour ouvrir son classeur et afficher les valeurs en direct pendant l'acquisition.
    ///
    /// On utilise du <c>dynamic</c> (late-binding COM via <c>Type.GetTypeFromProgID("Excel.Application")</c>)
    /// plutôt que le package Microsoft.Office.Interop.Excel : cela évite la dépendance à la PIA
    /// <c>office.dll</c> qui n'est pas installée par défaut avec toutes les versions d'Office et
    /// qui plante l'assembly loader à l'exécution avec "Could not load file or assembly 'office'".
    ///
    /// Tous les accès COM passent par <see cref="Marshal.ReleaseComObject"/> pour éviter qu'Excel
    /// ne reste en tâche de fond après la fermeture de Metrologo (zombie process).
    /// </summary>
    public sealed class ExcelInteropHost : IDisposable
    {
        private static readonly Lazy<ExcelInteropHost> _instance = new(() => new ExcelInteropHost());
        public static ExcelInteropHost Instance => _instance.Value;

        private dynamic? _excel;
        private dynamic? _classeurActif;
        private dynamic? _feuilleMesure;
        private string _cheminClasseurActif = string.Empty;

        /// <summary>
        /// Chemin du Metrologo.xla pré-ouvert dans l'instance Excel (utilisé pour résoudre
        /// les formules <c>[1]!Cal_xxx(...)</c> des rapports). Stocké pour exempter le .xla
        /// de <see cref="FermerClasseursParasitesInterne"/> qui sinon le refermerait au
        /// premier nettoyage. Vide tant que le .xla n'a pas été pré-chargé.
        /// </summary>
        private string _cheminXlaPreCharge = string.Empty;

        /// <summary>
        /// Sigma relatif (n-1) par numéro de gate accumulé pendant une session de mesure
        /// Stabilité. Utilisé à la dernière gate pour calibrer l'axe Y log du graphe Stab
        /// (<see cref="CalibrerAxeYGrapheStabViaComInterne"/>). Reset à chaque nouvelle
        /// session Stab via <see cref="ReinitialiserSigmasStab"/>.
        /// </summary>
        private readonly Dictionary<int, double> _sigmasRelatifsParGate = new();

        public void ReinitialiserSigmasStab() => _sigmasRelatifsParGate.Clear();

        /// <summary>
        /// Liste les processus EXCEL.EXE actuellement en cours sur la machine, EXCLUANT notre
        /// instance Excel cachée (identifiée par <see cref="_excelPid"/>). Utilisé pour détecter
        /// les classeurs Excel ouverts par l'utilisateur en parallèle de Metrologo — qui peuvent
        /// verrouiller le fichier .xlsx de la mesure et faire échouer ClosedXML.
        ///
        /// Retourne une liste vide si seul notre Excel COM tourne (cas normal) ou si Excel n'est
        /// pas installé.
        /// </summary>
        public List<Process> ListerExcelsExternes()
        {
            var liste = new List<Process>();
            try
            {
                foreach (var p in Process.GetProcessesByName("EXCEL"))
                {
                    try
                    {
                        if (p.Id != _excelPid && !p.HasExited)
                        {
                            liste.Add(p);
                        }
                    }
                    catch { /* process déjà mort entre Get et test */ }
                }
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "LISTER_EXCELS_EXTERNES_KO",
                    $"Énumération des process EXCEL échouée : {ex.Message}");
            }
            return liste;
        }

        /// <summary>
        /// Termine tous les processus EXCEL.EXE externes (Process.Kill) — c'est-à-dire ceux qui
        /// ne sont pas notre instance Excel COM cachée. À appeler après confirmation utilisateur
        /// pour libérer un fichier .xlsx verrouillé par une fenêtre Excel ouverte en parallèle.
        ///
        /// Retourne le nombre de process effectivement fermés. Best-effort : si un kill échoue
        /// (permissions, race condition), on continue avec les suivants.
        /// </summary>
        public int FermerExcelsExternes() => FermerExcels(ListerExcelsExternes());

        /// <summary>
        /// Termine la liste de processus EXCEL.EXE fournie (Process.Kill). Best-effort :
        /// un kill qui échoue (permissions, race) n'interrompt pas les suivants.
        /// Retourne le nombre de process effectivement fermés.
        /// </summary>
        public int FermerExcels(IEnumerable<Process> procs)
        {
            int nbFermes = 0;
            foreach (var p in procs)
            {
                try
                {
                    int pid = p.Id;
                    p.Kill();
                    p.WaitForExit(3000);
                    nbFermes++;
                    JournalLog.Info(CategorieLog.Excel, "EXCEL_EXTERNE_FERME",
                        $"Process EXCEL.EXE externe (PID={pid}) terminé via Process.Kill.");
                }
                catch (Exception ex)
                {
                    JournalLog.Warn(CategorieLog.Excel, "EXCEL_EXTERNE_KILL_KO",
                        $"Kill du process EXCEL.EXE (PID={p.Id}) échoué : {ex.Message}");
                }
            }
            return nbFermes;
        }

        /// <summary>
        /// Classe les EXCEL.EXE externes en deux catégories :
        ///   - <c>visibles</c> : possèdent une fenêtre principale (MainWindowHandle != 0) =
        ///     un vrai classeur ouvert par l'utilisateur, susceptible de contenir un travail
        ///     non sauvegardé → ne JAMAIS fermer sans confirmation.
        ///   - <c>fantomes</c> : aucun fenêtre = reliquat de pilotage COM (session Metrologo
        ///     précédente, complément, crash). Non visible par l'utilisateur, donc fermable
        ///     en silence pour libérer un éventuel verrou de fichier.
        /// </summary>
        /// <summary>
        /// Photographie l'ensemble des PID EXCEL.EXE actuellement en cours. À appeler
        /// JUSTE avant <c>Activator.CreateInstance("Excel.Application")</c> pour pouvoir,
        /// après création, identifier de façon fiable le PID de NOTRE nouvelle instance
        /// (cf. <see cref="ResoudrePidExcelCree"/>). Plus robuste que la lecture du HWND,
        /// qui échoue quand l'instance cachée n'a pas (encore) de fenêtre exploitable.
        /// </summary>
        private static HashSet<int> SnapshotPidsExcel()
        {
            var set = new HashSet<int>();
            try
            {
                foreach (var p in Process.GetProcessesByName("EXCEL"))
                {
                    try { set.Add(p.Id); } catch { /* process mourant */ }
                }
            }
            catch { /* énumération impossible : on retournera un set vide */ }
            return set;
        }

        /// <summary>
        /// Renvoie le PID du process EXCEL.EXE apparu depuis <paramref name="avant"/>
        /// (le snapshot pris avant <c>CreateInstance</c>) — c'est notre instance COM.
        /// Retourne 0 si rien de neuf (timing/échec), auquel cas l'appelant retombe sur
        /// la lecture du HWND comme solution de repli.
        /// </summary>
        private static int ResoudrePidExcelCree(HashSet<int> avant)
        {
            try
            {
                foreach (var p in Process.GetProcessesByName("EXCEL"))
                {
                    try { if (!avant.Contains(p.Id)) return p.Id; } catch { }
                }
            }
            catch { }
            return 0;
        }

        public (List<Process> visibles, List<Process> fantomes) ListerExcelsExternesClasses()
        {
            var visibles = new List<Process>();
            var fantomes = new List<Process>();
            foreach (var p in ListerExcelsExternes())
            {
                try
                {
                    if (p.MainWindowHandle != IntPtr.Zero) visibles.Add(p);
                    else fantomes.Add(p);
                }
                catch { fantomes.Add(p); /* process mourant : traité comme fantôme */ }
            }
            return (visibles, fantomes);
        }

        /// <summary>
        /// Vrai si un classeur de mesure est actuellement ouvert dans l'instance Excel.
        /// Utilisé par <see cref="MesureOrchestrator"/> pour décider du chemin :
        ///   - false (= 1ère mesure de la session app) → ClosedXML init + OuvrirEtAfficher
        ///   - true (= mesure 2+ sur même FI ou multi-gates) → <see cref="AjouterFeuilleMesureAsync"/>
        ///     en pur COM sans jamais fermer le classeur (élimine la fenêtre grise SDI).
        /// </summary>
        public bool AClasseurActif
        {
            get { lock (_sync) { return _classeurActif != null; } }
        }

        /// <summary>Chemin du classeur actuellement ouvert dans Excel (vide si aucun).</summary>
        public string CheminClasseurActif
        {
            get { lock (_sync) { return _cheminClasseurActif; } }
        }

        /// <summary>
        /// Désactive le rendu Excel (ScreenUpdating=false) le temps d'une transition
        /// fermeture/réouverture du classeur. Évite que la fenêtre Excel ne s'affiche
        /// brièvement vide (shell SDI grise) entre deux mesures. À appeler avant
        /// <c>FermerClasseurActifAsync</c> + init ClosedXML ; le rendu est restauré
        /// automatiquement par <c>OuvrirEtAfficherAsync</c> en fin de transition.
        /// </summary>
        public Task GelerAffichageAsync() => Task.Run(() =>
        {
            lock (_sync)
            {
                if (_excel == null) return;
                try { _excel.ScreenUpdating = false; } catch { /* best-effort */ }
            }
        });


        /// <summary>
        /// Déplace toutes les fenêtres top-level Excel du process hôte hors de l'écran visible
        /// (-32000, -32000) en mémorisant leur position pour restauration ultérieure.
        /// Utilisé pendant la transition fermeture/réouverture du Workbook entre 2 mesures :
        /// évite que la « shell SDI grise » apparaisse à l'écran pendant l'init ClosedXML,
        /// sans avoir à toggler Application.Visible (qui causait clignotement + perte de focus).
        /// </summary>
        public Task DeplacerExcelHorsEcranAsync() => Task.Run(() =>
        {
            lock (_sync)
            {
                if (_excel == null || _excelPid <= 0) return;
                _positionsFenetresExcelSauvees.Clear();
                try
                {
                    var positions = new Dictionary<IntPtr, (int, int)>();
                    EnumWindows((hWnd, _) =>
                    {
                        try
                        {
                            GetWindowThreadProcessId(hWnd, out uint pid);
                            if (pid != (uint)_excelPid) return true;
                            if (GetParent(hWnd) != IntPtr.Zero) return true;   // pas top-level
                            if (!IsWindowVisible(hWnd)) return true;
                            if (GetWindowRect(hWnd, out RECT r))
                            {
                                positions[hWnd] = (r.Left, r.Top);
                                SetWindowPos(hWnd, IntPtr.Zero, -32000, -32000, 0, 0,
                                    SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
                            }
                        }
                        catch { }
                        return true;
                    }, IntPtr.Zero);
                    foreach (var kv in positions) _positionsFenetresExcelSauvees[kv.Key] = kv.Value;
                }
                catch (Exception ex)
                {
                    JournalLog.Warn(CategorieLog.Excel, "EXCEL_OFFSCREEN_KO",
                        $"Déplacement Excel hors écran échoué : {ex.Message}");
                }
            }
        });

        /// <summary>
        /// Restaure les fenêtres Excel à leur position d'origine après une transition
        /// hors écran. Si une fenêtre du process n'était pas mémorisée (créée pendant la
        /// transition, ex: nouveau Workbook), on la place à (100, 100) par défaut.
        /// </summary>
        public Task RemettreExcelEnEcranAsync() => Task.Run(() =>
        {
            lock (_sync)
            {
                if (_excel == null || _excelPid <= 0) return;
                try
                {
                    var anciennesPositions = new Dictionary<IntPtr, (int X, int Y)>(_positionsFenetresExcelSauvees);
                    EnumWindows((hWnd, _) =>
                    {
                        try
                        {
                            GetWindowThreadProcessId(hWnd, out uint pid);
                            if (pid != (uint)_excelPid) return true;
                            if (GetParent(hWnd) != IntPtr.Zero) return true;
                            if (!IsWindowVisible(hWnd)) return true;
                            // Si on connaît sa position d'origine, on l'y remet. Sinon défaut.
                            if (anciennesPositions.TryGetValue(hWnd, out var pos))
                            {
                                SetWindowPos(hWnd, IntPtr.Zero, pos.X, pos.Y, 0, 0,
                                    SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
                            }
                            else
                            {
                                // Nouvelle fenêtre apparue pendant la transition (ex: Workbook
                                // ouvert après Open). Position d'écran par défaut.
                                SetWindowPos(hWnd, IntPtr.Zero, 100, 100, 0, 0,
                                    SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
                            }
                        }
                        catch { }
                        return true;
                    }, IntPtr.Zero);
                    _positionsFenetresExcelSauvees.Clear();
                }
                catch (Exception ex)
                {
                    JournalLog.Warn(CategorieLog.Excel, "EXCEL_ONSCREEN_KO",
                        $"Restauration position Excel échouée : {ex.Message}");
                }
            }
        });

        /// <summary>
        /// PID du process EXCEL.EXE qu'on a démarré, pour pouvoir le tuer en urgence
        /// si les COM calls hangent (fermeture manuelle Excel par l'utilisateur pendant
        /// une écriture en cours = appels RPC bloqués jusqu'à timeout système ~60 s).
        /// 0 = pas démarré.
        /// </summary>
        private int _excelPid;

        private readonly object _sync = new();
        private bool _disposed;

        // P/Invoke pour récupérer le PID du process Excel à partir de son handle.
        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // P/Invoke pour énumérer/masquer les fenêtres top-level Excel parasites (la fenêtre
        // « shell » d'Excel SDI qui reste visible et grise à côté du rapport).
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool BringWindowToTop(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }
        private const int SW_HIDE = 0;
        private const int SW_MINIMIZE = 6;
        private const int SW_FORCEMINIMIZE = 11;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOACTIVATE = 0x0010;
        private static readonly IntPtr HWND_BOTTOM = new IntPtr(1);

        // Mémorise la position originale de la fenêtre Excel hôte avant déplacement
        // hors écran (transition entre 2 mesures). Restaurée après la réouverture.
        private readonly Dictionary<IntPtr, (int X, int Y)> _positionsFenetresExcelSauvees = new();

        // Timer de masquage continu des fenêtres shell : se déclenche toutes les 500 ms tant
        // qu'un classeur est ouvert dans Excel. Garantit que toute shell qui apparaîtrait
        // (1ère ouverture, race COM, etc.) est masquée dans la demi-seconde max — même si la
        // tentative initiale a échoué pour cause de timing.
        private System.Threading.Timer? _timerMasquageShell;
        private readonly object _timerLock = new();

        private ExcelInteropHost() { }

        /// <summary>Vrai si Excel est démarré et prêt à ouvrir un classeur.</summary>
        public bool EstDemarre => _excel != null;

        /// <summary>
        /// Vérifie au niveau OS si le process Excel (PID enregistré) est encore vivant.
        /// Ne prend PAS le lock <c>_sync</c> — peut être appelé en boucle par un watchdog
        /// sans bloquer ni être bloqué par les COM calls en cours. Retourne false aussi
        /// si le PID n'est pas connu (Excel jamais démarré).
        /// </summary>
        public bool EstProcessExcelEnVie()
        {
            int pid = _excelPid;
            if (pid <= 0) return false;
            try
            {
                var p = Process.GetProcessById(pid);
                return !p.HasExited;
            }
            catch (ArgumentException) { return false; }   // process introuvable = mort
            catch { return true; }   // erreur transitoire, on suppose vivant (best-effort)
        }

        /// <summary>
        /// Démarre une instance Excel cachée. À appeler au lancement de l'application pour
        /// payer le coût d'initialisation COM pendant que l'utilisateur configure sa mesure.
        /// </summary>
        public Task DemarrerAsync() => Task.Run(() =>
        {
            lock (_sync)
            {
                if (_excel != null) return;
                try
                {
                    var excelType = Type.GetTypeFromProgID("Excel.Application");
                    if (excelType == null)
                    {
                        JournalLog.Warn(CategorieLog.Excel, "EXCEL_HOTE_ABSENT",
                            "Excel n'est pas installé sur ce poste : l'affichage live des mesures "
                            + "sera indisponible (les fichiers seront quand même sauvegardés).");
                        return;
                    }

                    // Photo des EXCEL.EXE déjà présents AVANT de créer le nôtre, pour
                    // identifier ensuite notre PID de façon fiable (cf. plus bas).
                    var pidsAvant = SnapshotPidsExcel();

                    _excel = Activator.CreateInstance(excelType);
                    if (_excel == null) return;

                    _excel.Visible = false;
                    _excel.DisplayAlerts = false;
                    _excel.ScreenUpdating = false;
                    _excel.AskToUpdateLinks = false;

                    // Identifie le PID de NOTRE instance par diff de la liste des process
                    // (fiable même sans fenêtre). Repli sur le HWND si le diff ne trouve rien.
                    // Ce PID sert à s'auto-exclure de la détection « Excel externe » et à
                    // TuerProcessExcelAsync() en cas de COM bloqué pendant une mesure.
                    _excelPid = ResoudrePidExcelCree(pidsAvant);
                    if (_excelPid == 0)
                    {
                        try
                        {
                            IntPtr hwnd = new IntPtr((int)_excel.Hwnd);
                            if (hwnd != IntPtr.Zero)
                            {
                                GetWindowThreadProcessId(hwnd, out uint pid);
                                _excelPid = (int)pid;
                            }
                        }
                        catch { /* best-effort, on saura par le journal si la mesure freeze */ }
                    }

                    JournalLog.Info(CategorieLog.Excel, "EXCEL_HOTE_DEMARRE",
                        $"Instance Excel cachée prête en arrière-plan (PID={_excelPid}).");

                    // Pré-ouvre Metrologo.xla comme classeur dans l'instance — c'est ce que
                    // faisait le Delphi historique. Sans ça, les formules [1]!Cal_xxx(...)
                    // des rapports affichent #NOM? jusqu'à ce que l'utilisateur clique
                    // « Ouvrir le classeur » dans le panneau « Liaisons de classeur ».
                    // Avec le .xla déjà dans Workbooks, Excel résout le link en silence.
                    PreChargerXlaInterne();
                }
                catch (Exception ex)
                {
                    JournalLog.Warn(CategorieLog.Excel, "EXCEL_HOTE_ERREUR",
                        $"Impossible de démarrer Excel en arrière-plan : {ex.Message}. "
                        + "Les mesures fonctionneront toujours, mais sans affichage live.",
                        new { ex.GetType().Name });
                    _excel = null;
                }
            }
        });

        /// <summary>
        /// Ouvre le classeur sauvegardé par <c>ExcelService</c> dans l'instance Excel cachée,
        /// positionne la feuille de mesure active, puis affiche la fenêtre Excel à l'utilisateur.
        /// Si un classeur était déjà ouvert, il est fermé en premier (choix : une seule mesure
        /// en direct à la fois — simplifie le suivi utilisateur).
        /// </summary>
        public Task OuvrirEtAfficherAsync(string cheminFichier, string nomFeuilleMesure)
            => Task.Run(() =>
            {
                lock (_sync)
                {
                    // Vérifie que l'instance Excel est encore vivante (l'utilisateur a pu
                    // fermer la fenêtre Excel manuellement, ou Excel a planté). Si KO,
                    // on relance une nouvelle instance avant de continuer — sinon RPC
                    // disconnecté (0x800706BA) au prochain accès à _excel.
                    if (!EstInstanceVivante())
                    {
                        JournalLog.Warn(CategorieLog.Excel, "EXCEL_HOTE_REDEMARRE",
                            "Instance Excel hôte indisponible (fermée ou crashée) — redémarrage automatique.");
                        RedemarrerInstanceInterne();
                        if (_excel == null) return;   // redémarrage a échoué, on abandonne
                    }

                    // Ferme tout classeur précédent (une seule mesure live à la fois).
                    FermerClasseurActifInterne();

                    // Workbooks.Open(Filename, UpdateLinks=0). On ne met PAS à jour les liens
                    // ici : le Metrologo.xla est déjà pré-ouvert dans l'instance Excel par
                    // PreChargerXlaInterne() au démarrage, donc Excel résout [1]!Cal_xxx(...)
                    // directement via le Workbook déjà en mémoire — pas besoin de re-scanner
                    // les liens (UpdateLinks=3 forcerait une recherche réseau lente et inutile).
                    _classeurActif = _excel.Workbooks.Open(cheminFichier, 0);
                    _cheminClasseurActif = cheminFichier;

                    // Active la feuille de mesure (ModFeuille dupliquée en Freq1/Stab1/...)
                    foreach (dynamic ws in _classeurActif.Worksheets)
                    {
                        if (string.Equals((string)ws.Name, nomFeuilleMesure, StringComparison.OrdinalIgnoreCase))
                        {
                            _feuilleMesure = ws;
                            ws.Activate();
                            break;
                        }
                    }

                    // Force Excel à recalculer formules + caches des graphes. Sans ça, le
                    // graphe Stab (axe Y log) reste vide : ClosedXML n'a pas peuplé le
                    // numCache des charts au moment de l'écriture, et l'ouverture avec
                    // UpdateLinks=0 ne déclenche pas le recalcul automatique. Le rebuild
                    // recompose la chaîne de calcul + régénère les caches des graphes.
                    try { _excel.CalculateFullRebuild(); }
                    catch (Exception ex)
                    {
                        JournalLog.Warn(CategorieLog.Excel, "EXCEL_RECALC_KO",
                            $"CalculateFullRebuild échoué : {ex.Message}");
                    }

                    // Ferme les éventuels classeurs résiduels (Book1 vide par défaut au
                    // démarrage Excel, ou anciens fichiers ouverts manuellement par
                    // l'utilisateur via double-clic) — sans ce nettoyage, plusieurs fenêtres
                    // Excel grises parasites apparaissent à côté du classeur actif après
                    // une relance.
                    FermerClasseursParasitesInterne();

                    _excel.ScreenUpdating = true;
                    _excel.Visible = true;

                    // Élargit la zone des onglets en bas (par défaut ~60% de la barre, ce qui
                    // tronque les onglets stab1..stab10 et oblige à naviguer avec les flèches).
                    // 0.85 = 85% de largeur réservée aux onglets, 15% restant pour la scrollbar
                    // horizontale → tous les onglets visibles d'un coup pour la plupart des FI.
                    try
                    {
                        dynamic w = _classeurActif.Windows[1];
                        try { w.TabRatio = 0.85; }
                        finally { try { Marshal.ReleaseComObject(w); } catch { } }
                    }
                    catch { /* best-effort, non bloquant */ }

                    // Masquage de la shell SDI vide qui apparaît avec Visible=true.
                    // Léger clignotement résiduel possible (le temps que ShowWindow s'applique).
                    MasquerFenetresShellExcelInterne();
                    MasquerFenetresNonActivesInterne();

                    // Force le rapport au premier plan (annule toute shell TopMost que Workbooks.Open
                    // aurait pu créer par-dessus).
                    ForcerRapportAuPremierPlanInterne();

                    // Démarre le timer de masquage continu : si une shell réapparaît plus tard
                    // (sur un Save, Open, ou autre opération COM qui réactive le window manager
                    // Excel), elle sera re-masquée dans la demi-seconde max.
                    DemarrerTimerMasquageShellInterne();

                    JournalLog.Info(CategorieLog.Excel, "EXCEL_CLASSEUR_OUVERT",
                        $"Classeur ouvert dans Excel : {Path.GetFileName(cheminFichier)} → feuille {nomFeuilleMesure}.");
                }
            });

        /// <summary>
        /// Écrit une valeur dans la ligne correspondant à la i-ème mesure. Les formules
        /// <c>Fréq. Réelle</c> (col C) et <c>F(i)-F(i+1)</c> (col D) sont déjà en place grâce
        /// à la phase d'initialisation ClosedXML.
        /// </summary>
        /// <param name="indexMesure">0-based — la 1ère mesure va en ligne <paramref name="ligneDebut"/>.</param>
        /// <param name="ligneDebut">Ligne Excel où commence la zone de mesures (9 dans le template).</param>
        public Task EcrireValeurLiveAsync(int indexMesure, double valeur, DateTime horodatage, int ligneDebut = 9)
            => Task.Run(() =>
            {
                lock (_sync)
                {
                    if (_feuilleMesure == null) return;
                    dynamic? range = null;
                    try
                    {
                        int row = ligneDebut + indexMesure;
                        // Tableau en A-D : col=1 (A) = HEURE, col=2 (B) = mesure brute.
                        //
                        // Écriture en bloc A:B via un Range + matrice [1,2] au lieu de 2 appels
                        // .Cells[r,c].Value2 successifs. Excel ne ré-active alors qu'une seule
                        // fois le rectangle de sélection visible sur la zone écrite (au lieu de
                        // 2 fois — col A puis col B). Réduit le clignotement / effet « valeur
                        // barrée pendant la saisie » côté utilisateur.
                        // Bonus perf : ~30-40 % plus rapide (1 marshal COM au lieu de 2).
                        range = _feuilleMesure.Range[
                            _feuilleMesure.Cells[row, 1],
                            _feuilleMesure.Cells[row, 2]];
                        range.Value2 = new object[1, 2]
                        {
                            { horodatage.ToString("HH:mm:ss"), valeur }
                        };
                    }
                    catch (Exception ex)
                    {
                        // Écriture live best-effort : une cellule plantée ne doit pas tuer la mesure.
                        JournalLog.Warn(CategorieLog.Excel, "EXCEL_ECRITURE_LIVE_ERREUR",
                            $"Écriture live impossible (i={indexMesure}) : {ex.Message}");
                    }
                    finally
                    {
                        if (range != null) { try { Marshal.ReleaseComObject(range); } catch { } }
                    }
                }
            });

        /// <summary>
        /// Écrit toutes les mesures d'une gate en un seul appel COM (Range.Value2 = matrix).
        /// Beaucoup plus rapide que N appels <see cref="EcrireValeurLiveAsync"/> quand l'utilisateur
        /// n'a pas besoin de voir les valeurs apparaître au fil de l'acquisition (ex: balayage de
        /// stabilité où chaque gate fait 30 mesures à ~10 ms — la boucle complète prend ~1 s,
        /// l'œil humain ne peut pas suivre, autant tout écrire d'un coup à la fin).
        /// </summary>
        /// <param name="ligneDebut">Ligne Excel où commence la zone (9 dans le template).</param>
        /// <param name="mesures">Liste ordonnée (timestamp, valeur). Index 0 = ligneDebut.</param>
        public Task EcrireValeursEnBlocAsync(int ligneDebut, IList<(DateTime ts, double valeur)> mesures)
            => Task.Run(() =>
            {
                lock (_sync)
                {
                    var feuille = _feuilleMesure;
                    if (feuille == null || mesures.Count == 0) return;
                    try
                    {
                        // Construction de la matrice 2D : N lignes × 2 colonnes (HEURE, mesure).
                        // Excel COM accepte object[,] indexé en 0-based pour Range.Value2.
                        object[,] matrix = new object[mesures.Count, 2];
                        for (int i = 0; i < mesures.Count; i++)
                        {
                            matrix[i, 0] = mesures[i].ts.ToString("HH:mm:ss");
                            matrix[i, 1] = mesures[i].valeur;
                        }

                        // Tableau en A-D : col 1 (A) = HEURE, col 2 (B) = mesure brute.
                        int ligneFin = ligneDebut + mesures.Count - 1;
                        dynamic plage = feuille.Range[
                            feuille.Cells[ligneDebut, 1],
                            feuille.Cells[ligneFin, 2]];
                        plage.Value2 = matrix;
                        Marshal.ReleaseComObject(plage);
                    }
                    catch (Exception ex)
                    {
                        JournalLog.Warn(CategorieLog.Excel, "EXCEL_BLOC_ECRITURE_ERREUR",
                            $"Écriture en bloc impossible ({mesures.Count} valeurs) : {ex.Message}");
                    }
                }
            });

        /// <summary>Sauvegarde le classeur actif sans le fermer (l'utilisateur le garde ouvert).</summary>
        public Task SauvegarderAsync() => Task.Run(() =>
        {
            lock (_sync)
            {
                if (_classeurActif == null) return;
                try { _classeurActif.Save(); }
                catch (Exception ex)
                {
                    JournalLog.Warn(CategorieLog.Excel, "EXCEL_SAUVEGARDE_ERREUR",
                        $"Échec de la sauvegarde via Interop : {ex.Message}");
                }
            }
        });

        /// <summary>
        /// Écrit une valeur dans une zone nommée (scope feuille) du classeur actif. Utilisé
        /// pour les saisies post-mesure (fréquence lue, incertitudes) qui complètent le rapport
        /// après acquisition des valeurs brutes.
        /// </summary>
        public Task EcrireZoneNommeeAsync(string zoneNom, object valeur) => Task.Run(() =>
        {
            lock (_sync)
            {
                if (_feuilleMesure == null) return;
                try
                {
                    // Names(name).RefersToRange.Value2 = valeur — sheet-scope names only.
                    dynamic nom = _feuilleMesure.Names.Item(zoneNom);
                    nom.RefersToRange.Value2 = valeur;
                    Marshal.ReleaseComObject(nom);
                }
                catch (Exception ex)
                {
                    // Zone absente ou erreur COM : on loggue en Warn, pas d'interruption.
                    JournalLog.Warn(CategorieLog.Excel, "EXCEL_ZONE_ECRITURE_ERREUR",
                        $"Impossible d'écrire la zone nommée « {zoneNom} » : {ex.Message}");
                }
            }
        });

        /// <summary>
        /// Ferme le classeur actif (avec sauvegarde) sans quitter Excel. À appeler quand on doit
        /// rouvrir le fichier avec un autre outil (ex : post-traitement ClosedXML).
        /// </summary>
        public Task<string> FermerClasseurActifAsync() => Task.Run(() =>
        {
            lock (_sync)
            {
                string chemin = _cheminClasseurActif;
                FermerClasseurActifInterne();
                return chemin;
            }
        });

        /// <summary>
        /// SOLUTION NUCLÉAIRE : tue le process EXCEL.EXE par son PID, sans passer par
        /// <c>_sync</c> ni par aucun appel COM. À utiliser uniquement en cas d'urgence
        /// quand un appel COM en cours est bloqué (typiquement : l'utilisateur a fermé
        /// la fenêtre Excel pendant qu'on écrivait une cellule → RPC en attente d'une
        /// réponse qui ne viendra jamais → timeout système ~60 s).
        /// <para/>
        /// Effets :
        ///   • Tous les appels COM en cours sur ce process lèvent immédiatement
        ///     <c>COMException 0x800706BA</c> (RPC server unavailable), les locks
        ///     <c>_sync</c> se libèrent, la mesure peut se terminer proprement.
        ///   • L'instance Excel partagée doit être recréée avant la prochaine mesure
        ///     (cf. <see cref="EstInstanceVivante"/> + <see cref="RedemarrerInstanceInterne"/>).
        /// <para/>
        /// Non bloquant : exécuté en Task.Run, ne nécessite pas de réponse.
        /// </summary>
        public Task TuerProcessExcelAsync() => Task.Run(() =>
        {
            // VOLONTAIREMENT pas de lock(_sync) ici — l'objectif est précisément de
            // débloquer les threads qui tiennent ce lock via COM hang. Si on prenait
            // le lock, on serait nous-même bloqué et on n'aiderait personne.
            int pid = _excelPid;
            if (pid <= 0)
            {
                JournalLog.Warn(CategorieLog.Excel, "EXCEL_KILL_PID_INCONNU",
                    "Tentative de tuer Excel mais PID inconnu — process probablement déjà mort.");
                return;
            }
            try
            {
                var proc = Process.GetProcessById(pid);
                proc.Kill(entireProcessTree: true);
                proc.WaitForExit(2000);
                JournalLog.Warn(CategorieLog.Excel, "EXCEL_KILL_OK",
                    $"Process Excel (PID={pid}) tué en urgence pour libérer les COM calls bloqués.");
            }
            catch (ArgumentException)
            {
                // Process déjà mort → c'est OK, c'est ce qu'on voulait.
                JournalLog.Info(CategorieLog.Excel, "EXCEL_KILL_DEJA_MORT",
                    $"Process Excel (PID={pid}) déjà terminé.");
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "EXCEL_KILL_ERREUR",
                    $"Kill du process Excel (PID={pid}) échoué : {ex.Message}");
            }
            finally
            {
                // Marque l'instance comme morte. Les locks tenus par les threads bloqués
                // vont se libérer dès que leur appel COM lève (immédiat après le kill).
                _excelPid = 0;
                _excel = null;
                _classeurActif = null;
                _feuilleMesure = null;
                _cheminClasseurActif = string.Empty;
            }
        });

        /// <summary>
        /// Ping l'instance Excel pour vérifier qu'elle répond encore au COM. Utilisé avant
        /// chaque ouverture de classeur pour éviter le RPC disconnect (0x800706BA) si
        /// l'utilisateur a fermé la fenêtre Excel manuellement ou si le process a crashé.
        /// </summary>
        private bool EstInstanceVivante()
        {
            if (_excel == null) return false;
            try
            {
                _ = _excel.Version;   // accès trivial qui plante si COM down
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Recrée une instance Excel après que la précédente ait été perdue (fermée par
        /// l'utilisateur, crashée, etc.). Doit être appelé sous lock <see cref="_sync"/>.
        /// </summary>
        private void RedemarrerInstanceInterne()
        {
            // Libère ce qui reste de l'ancienne instance (best-effort).
            try { if (_classeurActif != null) Marshal.ReleaseComObject(_classeurActif); } catch { }
            try { if (_feuilleMesure != null) Marshal.ReleaseComObject(_feuilleMesure); } catch { }
            try { if (_excel != null) Marshal.ReleaseComObject(_excel); } catch { }
            _classeurActif = null;
            _feuilleMesure = null;
            _excel = null;
            _cheminClasseurActif = string.Empty;
            _cheminXlaPreCharge = string.Empty;   // l'instance morte avait son propre .xla — invalide

            try
            {
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return;
                var pidsAvant = SnapshotPidsExcel();
                _excel = Activator.CreateInstance(excelType);
                if (_excel == null) return;
                _excel.Visible = false;
                _excel.DisplayAlerts = false;
                _excel.ScreenUpdating = false;
                _excel.AskToUpdateLinks = false;

                // Re-pré-ouvre Metrologo.xla pour que les formules [1]!Cal_xxx des prochains
                // rapports se résolvent silencieusement (cf. PreChargerXlaInterne).
                PreChargerXlaInterne();

                // Recapture le PID du NOUVEAU process Excel — sinon TuerProcessExcelAsync
                // taperait sur un PID obsolète et ne servirait à rien. Diff fiable + repli HWND.
                _excelPid = ResoudrePidExcelCree(pidsAvant);
                if (_excelPid == 0)
                {
                    try
                    {
                        IntPtr hwnd = new IntPtr((int)_excel.Hwnd);
                        if (hwnd != IntPtr.Zero)
                        {
                            GetWindowThreadProcessId(hwnd, out uint pid);
                            _excelPid = (int)pid;
                        }
                    }
                    catch { _excelPid = 0; }
                }
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "EXCEL_HOTE_REDEMARRAGE_KO",
                    $"Redémarrage Excel échoué : {ex.Message}");
                _excel = null;
                _excelPid = 0;
            }
        }

        /// <summary>
        /// Ouvre Metrologo.xla comme classeur dans l'instance Excel pour que les formules
        /// <c>[1]!Cal_xxx(...)</c> des rapports se résolvent silencieusement à l'ouverture.
        /// Reproduit ce que faisait le Delphi historique (cf. F_Main.pas: XLApp.Workbooks.Open(S_XLA, ...)).
        /// <para/>
        /// Sans ce pré-chargement, Excel affiche le panneau « Liaisons de classeur » qui force
        /// l'utilisateur à cliquer « Ouvrir le classeur » pour résoudre les fonctions VBA Cal_xxx.
        /// Un add-in chargé via HKCU\Excel\Options\OPENn ne suffit PAS : il fournit
        /// <c>=Cal_xxx(...)</c> direct mais pas <c>=[1]!Cal_xxx(...)</c> qui réfère un classeur.
        /// <para/>
        /// Doit être appelée sous lock <see cref="_sync"/> avec <see cref="_excel"/> non null.
        /// </summary>
        private void PreChargerXlaInterne()
        {
            if (_excel == null) return;
            string xlaPath = Preferences.CheminMacroXLA;
            if (string.IsNullOrWhiteSpace(xlaPath) || !File.Exists(xlaPath))
            {
                JournalLog.Warn(CategorieLog.Excel, "XLA_ABSENT",
                    $"Metrologo.xla introuvable à « {xlaPath} » — les cellules calculées des rapports afficheront #NOM?. "
                    + "Vérifier le chemin via Administration > Macro VBA.");
                return;
            }
            try
            {
                // UpdateLinks=0 : le .xla n'a pas de liens externes à mettre à jour ;
                // ReadOnly=true : sécurité, on ne risque pas d'altérer le .xla par accident.
                dynamic wb = _excel.Workbooks.Open(xlaPath, 0, true);
                _cheminXlaPreCharge = (string)wb.FullName;
                // Force IsAddin=true pour qu'Excel masque la fenêtre du .xla : sinon Excel 365
                // traite un .xla ouvert via Workbooks.Open comme un classeur normal et affiche
                // une fenêtre grise vide à côté du rapport quand l'utilisateur passe en Visible=true.
                // Une fois IsAddin=true, le classeur reste résolvable par les links mais invisible.
                try { wb.IsAddin = true; } catch { /* déjà true sur la plupart des .xla, on tente quand même */ }
                // On libère le proxy COM : Excel garde le classeur en mémoire via la collection
                // Workbooks, mais on n'a plus besoin de la référence côté C#.
                try { Marshal.ReleaseComObject(wb); } catch { }
                JournalLog.Info(CategorieLog.Excel, "XLA_PRE_CHARGE",
                    $"Metrologo.xla pré-ouvert dans Excel ({_cheminXlaPreCharge}) — formules [1]!Cal_xxx résolues auto.");
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "XLA_PRE_CHARGE_KO",
                    $"Pré-chargement de Metrologo.xla échoué : {ex.Message}. "
                    + "Les rapports afficheront #NOM? sur les cellules calculées.");
            }
        }

        /// <summary>
        /// Masque toutes les fenêtres Excel SAUF celle du classeur actif (rapport en cours).
        /// Excel a la fâcheuse habitude de réafficher les fenêtres du .xla (et de tout autre
        /// classeur résiduel) à chaque toggle <c>Visible=false→true</c> de l'instance hôte.
        /// Sans cette méthode : une 2e fenêtre grise vide apparaît à côté du rapport dès la
        /// 2e mesure (toggle qui réveille la fenêtre du .xla).
        /// Doit être appelée sous lock <see cref="_sync"/>, après <c>_excel.Visible=true</c>.
        /// </summary>
        private void MasquerFenetresNonActivesInterne()
        {
            if (_excel == null) return;
            try
            {
                var diag = new System.Text.StringBuilder();
                int nb = _excel.Windows.Count;
                diag.Append($"actif=[{_cheminClasseurActif}] nbWindows={nb} ");
                for (int i = 1; i <= nb; i++)
                {
                    dynamic w = _excel.Windows[i];
                    try
                    {
                        dynamic wb = w.Parent;
                        try
                        {
                            string fullName = (string)wb.FullName;
                            string caption = (string)w.Caption;
                            bool visAvant = (bool)w.Visible;
                            bool estClasseurActif = ChemPathEgal(fullName, _cheminClasseurActif);
                            string action = "skip";
                            if (!estClasseurActif)
                            {
                                try { w.Visible = false; action = "MASQUE"; } catch { action = "ECHEC"; }
                            }
                            diag.Append($"| W#{i}={caption}:visAv={visAvant}:wb=[{fullName}]:{action} ");
                        }
                        finally { try { Marshal.ReleaseComObject(wb); } catch { } }
                    }
                    catch { /* best-effort */ }
                    finally { try { Marshal.ReleaseComObject(w); } catch { } }
                }
                JournalLog.Info(CategorieLog.Excel, "EXCEL_MASQUE_DIAG", diag.ToString());
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "EXCEL_MASQUE_KO",
                    $"Re-masquage des fenêtres non-actives échoué : {ex.Message}");
            }
        }

        /// <summary>
        /// Ferme tous les classeurs ouverts dans l'instance Excel hôte SAUF
        /// <see cref="_classeurActif"/> ET <see cref="_cheminXlaPreCharge"/>. Évite que des
        /// fenêtres grises parasites (Book1 par défaut, anciens classeurs) restent visibles
        /// à côté du classeur de mesure, sans refermer le .xla qu'on veut garder en mémoire.
        /// Doit être appelée sous lock <see cref="_sync"/>.
        /// </summary>
        private void FermerClasseursParasitesInterne()
        {
            if (_excel == null) return;
            try
            {
                // On itère sur Workbooks via index décroissant — la collection se met à
                // jour à chaque Close, et un foreach risque de lever InvalidOperationException.
                int nb = _excel.Workbooks.Count;
                var diag = new System.Text.StringBuilder($"actif=[{_cheminClasseurActif}] xla=[{_cheminXlaPreCharge}] nb={nb} ");
                for (int i = nb; i >= 1; i--)
                {
                    dynamic wb = _excel.Workbooks[i];
                    try
                    {
                        string fullName = (string)wb.FullName;
                        // Comparaison robuste : Path.GetFullPath normalise les variations
                        // (backslash final, capitalisation drive, raccourcis 8.3, /vs\, etc.)
                        // qui font échouer un simple string.Equals — ce qui faisait fermer
                        // le rapport tout juste ouvert et donnait une fenêtre Excel grise
                        // vide à côté.
                        bool estClasseurActif = ChemPathEgal(fullName, _cheminClasseurActif);
                        bool estXlaPreCharge = !string.IsNullOrEmpty(_cheminXlaPreCharge)
                            && ChemPathEgal(fullName, _cheminXlaPreCharge);
                        string action = "skip";
                        if (!estClasseurActif && !estXlaPreCharge)
                        {
                            try { wb.Close(false); action = "FERME"; }
                            catch { action = "ECHEC_CLOSE"; }
                        }
                        diag.Append($"| WB#{i}=[{fullName}]:{action} ");
                    }
                    catch { /* best-effort */ }
                    finally
                    {
                        try { Marshal.ReleaseComObject(wb); } catch { }
                    }
                }
                JournalLog.Info(CategorieLog.Excel, "EXCEL_PARASITES_DIAG", diag.ToString());
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "EXCEL_PARASITES_KO",
                    $"Nettoyage des classeurs parasites échoué : {ex.Message}");
            }
        }

        // ============================================================================
        //  OPTION A — Finalisation de mesure 100% en COM
        // ============================================================================
        //  Au lieu du cycle ferme-COM / écrit-ClosedXML / rouvre-COM (qui cause la fenêtre
        //  grise à la fin de la mesure), on écrit directement dans le Workbook ouvert via
        //  Excel COM. Plus aucune fermeture/réouverture → l'instance reste stable, le
        //  rapport reste visible, plus de fenêtre Frame parasite.
        //
        //  Couvre : coeffs d'incertitude (ZNCoeffA/B/C/D) + ligne Récap (Freq) + Save.
        //  Limitations actuelles : Stab + Tachy fréquence auxiliaire couverts par le même
        //  chemin ; le graphe Stab reste créé via le chemin ClosedXML historique.
        // ============================================================================

        /// <summary>
        /// Finalise la mesure en écrivant via Excel COM (sans cycle ferme/rouvre) :
        /// les coefficients d'incertitude résolus depuis le module sélectionné, puis la
        /// ligne Récap correspondante au type de mesure, puis Save du Workbook actif.
        /// L'instance Excel et le rapport restent visibles tout du long.
        /// </summary>
        public Task FinaliserMesureViaComAsync(Mesure mesure, List<double> resultats,
            string nomFeuille, double tempsGateSecondes, bool isDerniereGate = false,
            int nbGatesStabBalayees = 0) => Task.Run(() =>
        {
            lock (_sync)
            {
                if (_excel == null || _classeurActif == null) return;

                // FREEZE TOTAL DU RENDU EXCEL pendant la finalisation. Sans ça, les ~80+
                // opérations COM sur la feuille Récap (Clear + Style + NumberFormat + Value2
                // + Formula + alignment H/V sur ~12 cellules) déclenchent chacune un repaint :
                // l'utilisateur voit Excel scroller, basculer entre feuilles, animer = effet
                // de "balayage" visible avant l'affichage final de Récap.
                bool screenUpdatingOriginal = true;
                int calculationOriginal = -4105;   // xlCalculationAutomatic
                try { screenUpdatingOriginal = (bool)_excel.ScreenUpdating; } catch { }
                try { calculationOriginal = (int)_excel.Calculation; } catch { }
                try { _excel.ScreenUpdating = false; } catch { }
                try { _excel.Calculation = -4135; /* xlCalculationManual */ } catch { }

                try
                {
                    // ⚠ CalculerEtEcrireVarianceViaComInterne RETIRÉ (avait été ajouté pour
                    //    la voie COM v2 qui rejetait la formule `=[1]!Cal_variance` en .Formula).
                    //    Avec le retour à la voie ClosedXML, ClosedXML pose correctement la formule
                    //    de Variance dans PreparerLignesMesureAsync — on doit la LAISSER en place,
                    //    pas l'écraser avec une valeur littérale. Sinon F43 (ZNVariance) perd sa
                    //    formule `[1]!Cal_variance(SUMSQ(G10:G38),ZNNbMesures,ZNFreqMoyReel)` qui
                    //    permet le recalcul live si l'utilisateur modifie les données.

                    // 1. Coeffs d'incertitude + sigma Stab si applicable
                    EcrireCoeffsIncertitudeViaComInterne(mesure, resultats, nomFeuille, tempsGateSecondes);

                    // 2. Ligne Récap (Freq ou Stab selon le type)
                    if (mesure.TypeMesure == TypeMesure.Frequence
                        || mesure.TypeMesure == TypeMesure.FreqAvantInterv
                        || mesure.TypeMesure == TypeMesure.FreqFinale
                        || mesure.TypeMesure == TypeMesure.Interval
                        || EnTetesMesureHelper.EstTachymetre(mesure.TypeMesure)
                        || mesure.TypeMesure == TypeMesure.Stroboscope)
                    {
                        AjouterLigneRecapFreqViaComInterne(mesure, nomFeuille);
                    }
                    else if (mesure.TypeMesure == TypeMesure.Stabilite)
                    {
                        AjouterLigneRecapStabViaComInterne(mesure, nomFeuille);

                        // À la dernière gate Stab : patcher le graphe via COM (plages des
                        // séries + bornes axe Y log) — équivalent de RendreGrapheStabDynamique
                        // + CalibrerAxeYGrapheStab de ExcelService, mais sans cycle ferme/rouvre.
                        if (isDerniereGate)
                        {
                            int nbGates = nbGatesStabBalayees > 0
                                ? nbGatesStabBalayees
                                : _sigmasRelatifsParGate.Count;
                            AjusterPlagesGrapheStabViaComInterne(nbGates);
                            CalibrerAxeYGrapheStabViaComInterne(_sigmasRelatifsParGate.Values);
                        }
                    }

                    // 3. Redirige le lien externe Metrologo.xla vers Preferences.CheminMacroXLA
                    //    via ChangeLink COM (équivalent du PatcherLienMacroXLA ClosedXML qui
                    //    n'est plus appelé en mode visible). Évite que Excel demande à pointer
                    //    le .xla si le rels contient des chemins obsolètes.
                    RedirigerLinksXlaViaComInterne();

                    // 4. ⚠ Recalc complet AVANT le Save. Toutes les formules ont été posées sous
                    //    Calculation=Manual (pas d'auto-recalc à chaque pose, pour éviter le
                    //    balayage Récap). Sans ce recalc, le fichier sauvegardé contient les
                    //    formules MAIS pas leurs valeurs caches → à la réouverture, les cellules
                    //    dépendant de liens externes (=[1]!Cal_variance, etc. → Metrologo.xla)
                    //    s'affichent comme "périmées" (style strikethrough Excel natif) car le
                    //    lien externe est désactivé par défaut et le cache est vide.
                    //
                    //    On force Calculation=Automatic puis Calculate() pour résoudre toutes
                    //    les formules, AVANT le Save → le .xlsx contient les valeurs à jour →
                    //    réouverture propre, pas de cellule barrée.
                    try { _excel.Calculation = -4105; /* xlCalculationAutomatic */ } catch { }
                    try { _classeurActif.Calculate(); } catch { }

                    // Configure le workbook pour ne PAS tenter d'actualiser les liens externes
                    // (vers Metrologo.xla) à la prochaine ouverture. Sans ce réglage, Excel
                    // recalcule les liens à chaque ouverture du .xlsx → marque le classeur
                    // "modifié" même si rien n'a été touché → propose enregistrer à la fermeture.
                    // xlUpdateLinksNever = 2 : pas de mise à jour auto, les valeurs caches du
                    // .xlsx (posées par notre Calculate ci-dessus) sont utilisées telles quelles.
                    try { _classeurActif.UpdateLinks = 2; /* xlUpdateLinksNever */ } catch { }

                    // 4.ter Verrouillage post-mesure : à la dernière gate, on re-protège toutes
                    //       les feuilles avec le mot de passe prédéfini AVANT le Save (pour que la
                    //       protection soit persistée). Le rapport devient non modifiable dans Excel
                    //       sans le mot de passe ; l'app le déprotège seule à la relance (mdp connu).
                    if (isDerniereGate)
                        ProtegerToutesFeuillesActifInterne();

                    // 5. Save (avec valeurs caches à jour grâce au Calculate ci-dessus)
                    try { _classeurActif.Save(); }
                    catch (Exception ex)
                    {
                        JournalLog.Warn(CategorieLog.Excel, "EXCEL_SAVE_COM_KO",
                            $"Save COM final échoué : {ex.Message}");
                    }

                    // Marque le classeur comme "non modifié" après le Save. Si Excel fait un
                    // recalc interne entre maintenant et la fermeture de notre instance hôte
                    // (ou même à la réouverture côté utilisateur), il considérera que le
                    // fichier sur disque correspond déjà à l'état mémoire → pas de prompt
                    // "Voulez-vous enregistrer les modifications ?".
                    try { _classeurActif.Saved = true; } catch { }

                    // 5. Active la feuille Récap. en fin de mesure (visibilité utilisateur) —
                    //    pour Stab on attend la dernière gate du balayage, sinon on resterait
                    //    sur Récap entre chaque gate au lieu de voir la feuille en cours.
                    bool peutBasculerRecap = mesure.TypeMesure != TypeMesure.Stabilite
                                              || isDerniereGate;
                    if (peutBasculerRecap)
                    {
                        try
                        {
                            dynamic recap = _classeurActif.Worksheets.Item("Récap.");
                            try { recap.Activate(); }
                            finally { try { Marshal.ReleaseComObject(recap); } catch { } }
                            JournalLog.Info(CategorieLog.Excel, "EXCEL_RECAP_ACTIVE",
                                "Feuille Récap. activée en fin de mesure (vue finale utilisateur).");
                        }
                        catch (Exception ex)
                        {
                            JournalLog.Warn(CategorieLog.Excel, "EXCEL_RECAP_ACTIVE_KO",
                                $"Activation Récap. échouée : {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    JournalLog.Erreur(CategorieLog.Excel, "FINALISER_VIA_COM_KO",
                        $"Finalisation mesure via COM échouée : {ex.Message}");
                }
                finally
                {
                    // Restaure les flags Excel dans l'ordre inverse :
                    //   1. Calculation = original   → repasse en Automatic (typiquement)
                    //   2. Calculate() FORCÉ        → Excel ne recalcule PAS automatiquement
                    //                                  les formules qui ont été posées pendant
                    //                                  la phase Manual. Sans ce force, les
                    //                                  cellules figées dans leur dernier état
                    //                                  (#VALEUR!, #DIV/0!, 0) gardent leur style
                    //                                  "erreur" du template (couleur grise +
                    //                                  strikethrough) même quand les données
                    //                                  saisies les auraient résolues.
                    //   3. ScreenUpdating = original → Excel dessine l'état final propre d'un
                    //                                  coup, formules à jour, sans le balayage
                    //                                  intermédiaire ni les cellules "barrées".
                    //   4. Saved = true              → marque le classeur "non modifié" APRÈS
                    //                                  toutes les opérations Excel (Calculate
                    //                                  inclus). Sans ce dernier flag, le
                    //                                  Calculate du finally re-marque le classeur
                    //                                  dirty même quand rien n'a vraiment changé,
                    //                                  ce qui fait que Excel propose d'enregistrer
                    //                                  à la fermeture côté utilisateur.
                    try { _excel.Calculation = calculationOriginal; } catch { }
                    try { _classeurActif.Calculate(); } catch { }
                    try { _excel.ScreenUpdating = screenUpdatingOriginal; } catch { }
                    try { _classeurActif.Saved = true; } catch { }
                }
            }
        });

        /// <summary>
        /// Ajoute une nouvelle feuille de mesure (copie de ModFeuille) dans le classeur
        /// actuellement ouvert, en pur COM, sans jamais fermer le Workbook. Remplace pour
        /// les mesures 2+ la séquence
        /// <c>FermerClasseurActif → ClosedXML InitialiserRapport + PreparerLignesMesure
        /// + Sauvegarder → OuvrirEtAfficher</c> qui laissait Excel <c>Visible=true</c>
        /// avec 0 Workbook visible pendant 1-3 secondes (= shell SDI grise visible).
        ///
        /// Reproduit exactement le travail de <c>ExcelService.InitialiserRapportAsync</c> +
        /// <c>PreparerLignesMesureAsync</c> :
        /// <list type="number">
        ///   <item><c>ModFeuille.Copy</c> en COM (place la copie en fin de classeur).</item>
        ///   <item>Rename, visibilité, déprotection, largeurs E/F/H.</item>
        ///   <item><b>Clonage des Names sheet-scope</b> : pour chaque Name workbook-scope dont le
        ///         RefersTo pointe vers <c>ModFeuille!</c>, création d'un Name sheet-scope sur la
        ///         nouvelle feuille (via <c>Worksheet.Names.Add</c>, pas <c>Workbook.Names.Add</c>) avec
        ///         le RefersTo redirigé vers la nouvelle feuille. <b>Sans ce clonage</b>, les formules
        ///         locales <c>ZNFreqMoyReel = AVERAGE(F9:F38)</c>, <c>ZNVariance</c>, etc. résoudraient
        ///         vers ModFeuille (vide) → variance/écart-type/incert KO (= cause du fail de la
        ///         précédente tentative option A).</item>
        ///   <item>En-têtes D/E/F/G + labels col E + colonnes n°Module/Fonction/Condition 1 ligne 9.</item>
        ///   <item>Colonne N conversion tr/min si tachymétrie.</item>
        ///   <item>Métadonnées via zones nommées sheet-scope (ZNNoFiche, ZNDate, ZNTypeMesure, …).</item>
        ///   <item>Coefficients hardcoded (1e-10 / 5e-13) — surcharge dans <c>FinaliserMesureViaComAsync</c>
        ///         si module d'incertitude sélectionné.</item>
        ///   <item>InsertRowsAbove + formules F (conversion) + G (delta) + N (tr/min) pour nbMesures &gt; 2.</item>
        ///   <item>Formules métier (ZNFreqMoyReel = AVERAGE, ZNVariance = Cal_variance VBA, etc.).</item>
        ///   <item>Activate la feuille (visible utilisateur immédiatement) + Save du Workbook.</item>
        /// </list>
        ///
        /// Le Workbook n'est jamais fermé entre le moment où il a été ouvert (mesure 1 via
        /// <see cref="OuvrirEtAfficherAsync"/>) et la fermeture définitive de l'app. Plus de
        /// moment 0-Workbook-visible → plus de shell SDI grise.
        /// </summary>
        /// <returns>Nom de la nouvelle feuille créée (ex: <c>"freq2"</c>, <c>"stab3"</c>).</returns>
        public Task<string> AjouterFeuilleMesureAsync(
            Mesure config, Rubidium rubidium, int gateInscrite)
            => Task.Run<string>(() =>
            {
                lock (_sync)
                {
                    if (_excel == null || _classeurActif == null)
                    {
                        throw new InvalidOperationException(
                            "AjouterFeuilleMesureAsync appelé sans classeur actif. "
                            + "Cette méthode ne doit être utilisée que pour les mesures 2+ "
                            + "(quand Excel a déjà un Workbook ouvert). Pour la 1ère mesure, "
                            + "utiliser ClosedXML.InitialiserRapportAsync + OuvrirEtAfficherAsync.");
                    }

                    dynamic? modFeuille = null;
                    dynamic? nouvelleFeuille = null;

                    // FREEZE COMPLET pendant la séquence COM (mesure : opération qui prend
                    // ~6 secondes sans freeze à cause des 33 Names.Add + 60 EcrireFormuleCellule
                    // + redraw à chaque écriture). Avec ScreenUpdating=false on tombe à ~1-2s.
                    // L'écran peut sembler "gris" pendant ce temps, mais le délai est tellement
                    // réduit que c'est imperceptible. ForcerRapportAuPremierPlanInterne à la fin
                    // force le rafraîchissement final.
                    bool screenUpdatingOriginal = true, enableEventsOriginal = true;
                    int calculationOriginal = -4105;
                    try { screenUpdatingOriginal = (bool)_excel.ScreenUpdating; } catch { }
                    try { enableEventsOriginal = (bool)_excel.EnableEvents; } catch { }
                    try { calculationOriginal = (int)_excel.Calculation; } catch { }
                    try { _excel.ScreenUpdating = false; } catch { }
                    try { _excel.EnableEvents = false; } catch { }
                    try { _excel.Calculation = -4135; /* xlCalculationManual */ } catch { }

                    try
                    {
                        // --- 1. Récupère ModFeuille (présent dans tous les templates) ---
                        try { modFeuille = _classeurActif.Worksheets.Item("ModFeuille"); }
                        catch (Exception ex)
                        {
                            throw new InvalidOperationException(
                                "ModFeuille introuvable dans le classeur actif — "
                                + "le fichier semble corrompu. Détail : " + ex.Message);
                        }

                        // --- 2. Génère un nom de feuille unique (freq2, stab3, etc.) ---
                        string nomNouvelleFeuille = TrouverNomFeuilleUniqueInterne(config.TypeMesure);

                        // Feuilles à nom fixe (avinter / fqfinale) : à la relance d'une mesure
                        // avant/après intervention sur la même FI, le nom existe déjà → on ÉCRASE
                        // l'ancienne feuille ET sa ligne Récap avant la copie (sinon le rename plus
                        // bas lèverait sur doublon et la Récap afficherait deux lignes). No-op pour
                        // les feuilles numérotées (nom toujours neuf).
                        SupprimerFeuilleEtRecapInterne(nomNouvelleFeuille);

                        // --- 3. Copie ModFeuille à la fin du classeur ---
                        // ⚠ PIÈGE CRITIQUE D'EXCEL COM ⚠ : ws.Copy(Before, After) — selon la doc
                        // Microsoft, "If neither Before nor After is specified, Excel creates a
                        // new workbook that contains the copied object." En late-binding dynamic,
                        // passer Type.Missing pour Before était interprété comme "non spécifié"
                        // → Excel créait un NOUVEAU WORKBOOK parasite à chaque appel !
                        // Solution : utiliser le paramètre nommé After: (C# dynamic supporte les
                        // named args via IDispatch). Excel comprend correctement et ajoute la
                        // feuille au classeur existant.
                        dynamic derniereFeuille = _classeurActif.Worksheets
                            [(int)_classeurActif.Worksheets.Count];
                        try
                        {
                            modFeuille.Copy(After: derniereFeuille);
                        }
                        finally { try { Marshal.ReleaseComObject(derniereFeuille); } catch { } }

                        // Après Copy, la feuille copiée devient ActiveSheet du classeur cible.
                        // On vérifie via _classeurActif.ActiveSheet (et pas Application.ActiveSheet
                        // qui pourrait pointer ailleurs si Excel a changé de classeur actif).
                        nouvelleFeuille = _classeurActif.ActiveSheet;

                        // --- 4. Rename + Visible + Unprotect ---
                        nouvelleFeuille.Name = nomNouvelleFeuille;
                        nouvelleFeuille.Visible = -1;   // xlSheetVisible
                        try { nouvelleFeuille.Unprotect("METROL"); }
                        catch
                        {
                            try { nouvelleFeuille.Unprotect("metrol"); }
                            catch { try { nouvelleFeuille.Unprotect(); } catch { } }
                        }

                        // Largeurs : le template a été décalé de 3 colonnes (E/F/H = ancien B/C/E).
                        try
                        {
                            ((dynamic)nouvelleFeuille.Columns["E"]).ColumnWidth = 30;
                            ((dynamic)nouvelleFeuille.Columns["F"]).ColumnWidth = 28;
                            ((dynamic)nouvelleFeuille.Columns["H"]).ColumnWidth = 30;
                        }
                        catch { /* best-effort */ }

                        // --- 5. CLONAGE DES NAMES SHEET-SCOPE (PARTIE CRITIQUE) ---
                        // Les templates ont les Names workbook-scope pointant vers ModFeuille!
                        // (33 sur METROLOGO.xlsx, idem sur Stab/Tachy). Worksheet.Copy COM ne
                        // les clone PAS automatiquement (il ne clone que les sheet-scope, qui
                        // n'existent pas sur ModFeuille). On doit donc recréer chaque Name en
                        // version sheet-scope sur la nouvelle feuille — sans ça, ZNFreqMoyReel,
                        // ZNVariance, ZNEcartType, ZNIncertEcartType, ZNFreqCorr résoudraient
                        // vers ModFeuille (vide) au lieu de la feuille de mesure.
                        int nbNamesClones = 0, nbNamesIgnores = 0;
                        var nomsAClones = new List<(string Nom, string RefersTo)>();
                        foreach (dynamic name in (dynamic)_classeurActif.Names)
                        {
                            try
                            {
                                string nomZone = (string)name.Name;
                                string refersTo = (string)name.RefersTo;
                                if (string.IsNullOrEmpty(refersTo)) continue;
                                if (!refersTo.Contains("ModFeuille!", StringComparison.OrdinalIgnoreCase))
                                    continue;
                                string newRefersTo = System.Text.RegularExpressions.Regex.Replace(
                                    refersTo, "ModFeuille!",
                                    $"'{nomNouvelleFeuille}'!",
                                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                                nomsAClones.Add((nomZone, newRefersTo));
                            }
                            catch { /* Name corrompu — skip silencieux */ }
                            finally { try { Marshal.ReleaseComObject(name); } catch { } }
                        }

                        foreach (var (nomZone, newRefersTo) in nomsAClones)
                        {
                            try
                            {
                                // Worksheet.Names.Add (sheet-scope) — PAS Workbook.Names.Add
                                // (workbook-scope, ce qui créerait un conflit avec l'existant).
                                ((dynamic)nouvelleFeuille.Names).Add(nomZone, newRefersTo);
                                nbNamesClones++;
                            }
                            catch (Exception)
                            {
                                nbNamesIgnores++;
                            }
                        }

                        JournalLog.Info(CategorieLog.Excel, "AJOUT_FEUILLE_NAMES_CLONES",
                            $"Clonage Names sheet-scope vers {nomNouvelleFeuille} : "
                            + $"{nbNamesClones} clonés, {nbNamesIgnores} ignorés "
                            + $"(sur {nomsAClones.Count} candidats trouvés pointant vers ModFeuille!).");

                        // --- 6. En-têtes adaptatifs col A/B/C/D + labels col B ---
                        var entetes = EnTetesMesureHelper.Pour(config.TypeMesure);
                        EcrireValeurCelluleInterne(nouvelleFeuille, "A7",  entetes.EnteteHeure);
                        EcrireValeurCelluleInterne(nouvelleFeuille, "B7",  entetes.EnteteMesuree);
                        EcrireValeurCelluleInterne(nouvelleFeuille, "C7",  entetes.EnteteReelle);
                        EcrireValeurCelluleInterne(nouvelleFeuille, "D7",  entetes.EnteteDelta);
                        EcrireValeurCelluleInterne(nouvelleFeuille, "B13", entetes.LabelMoyenne);
                        EcrireValeurCelluleInterne(nouvelleFeuille, "B21", entetes.LabelFreqRef);
                        EcrireValeurCelluleInterne(nouvelleFeuille, "B23", entetes.LabelFreqCorr);
                        EcrireValeurCelluleInterne(nouvelleFeuille, "B25", entetes.LabelIncertResol);
                        EcrireValeurCelluleInterne(nouvelleFeuille, "B31", entetes.LabelIncertGlob);

                        // --- 7. Module d'incertitude affiché UNE seule fois ---
                        // On écrit le module SÉLECTIONNÉ dans la config (config.NumModuleIncertitude),
                        // dans la cellule valeur à côté du label « Module = » : F10 (layout standard)
                        // ou G10 (tachy, 1 colonne de plus).
                        bool estTachy = EnTetesMesureHelper.EstTachymetre(config.TypeMesure);
                        string moduleSelectionne = config.NumModuleIncertitude ?? string.Empty;
                        string celluleModule = estTachy ? "G10" : "F10";
                        if (!string.IsNullOrEmpty(moduleSelectionne))
                            EcrireValeurCelluleInterne(nouvelleFeuille, celluleModule, moduleSelectionne);

                        // --- 8. Tachymétrie : col K (conversion Hz → tr/min) ---
                        if (estTachy)
                        {
                            EcrireValeurCelluleInterne(nouvelleFeuille, "K7", "Vitesse (tr/min)");
                            try { ((dynamic)nouvelleFeuille.Columns["K"]).ColumnWidth = 18; } catch { }
                            EcrireFormuleCelluleInterne(nouvelleFeuille, "K9",  "=B9*60");
                            EcrireFormuleCelluleInterne(nouvelleFeuille, "K10", "=B10*60");
                        }

                        // --- 9. Métadonnées via zones nommées sheet-scope (clonées plus haut) ---
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNNoFiche", config.NumFI);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNDate",
                            DateTime.Now.ToString("dd/MM/yyyy"));
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNTypeMesure",
                            EnTetesMesureHelper.LibelleType(config.TypeMesure));
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNFreqUtilise",
                            NomAppareilDepuisCatalogueLocal(config.IdModeleCatalogue));
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNRubidium",
                            rubidium.DesignationAvecRaccord);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNGate",
                            EnTetesMesureHelper.LibelleGate(gateInscrite));
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNLibGate",
                            EnTetesMesureHelper.LibelleGate(gateInscrite));
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNValGateSecondes",
                            EnTetesMesureHelper.SecondesGate(gateInscrite));
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNModeMesure",
                            config.ModeMesure == ModeMesure.Direct ? "Direct" : "Indirect");
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNCoeffMult",
                            (double)config.IndexMultiplicateur);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNValFNominale",
                            config.FNominale);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNNbMesures",
                            (double)config.NbMesures);
                        // Résolution = 0 à l'init : seule la saisie post-mesure de l'utilisateur
                        // peut la renseigner (cf. ExcelService.InitialiserRapportAsync pour le
                        // raisonnement complet).
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNIncertResol", 0.0);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNIncertSup",
                            config.IncertSupp);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNFreqRef",
                            rubidium.FrequenceMoyenne);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNCoeffA", 1e-10);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNCoeffB", 5e-13);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNNbMesAccredite", 30.0);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNTempsMesureAccredite", 10.0);

                        // --- 10. Insertion des lignes additionnelles (PreparerLignesMesure) ---
                        // Le template a 2 lignes pré-créées (9 et 10). Pour nbMesures > 2, on
                        // insère (nbMesures - 2) rangées AVANT le point d'insertion (par défaut
                        // ligne 11 = position de ZNPointInsertion).
                        int nbMesures = config.NbMesures;
                        if (nbMesures > 2)
                        {
                            int pointInsertionRow = 11;
                            try
                            {
                                dynamic nameIns = nouvelleFeuille.Names.Item("ZNPointInsertion");
                                try
                                {
                                    dynamic refRange = nameIns.RefersToRange;
                                    try { pointInsertionRow = (int)refRange.Row; }
                                    finally { try { Marshal.ReleaseComObject(refRange); } catch { } }
                                }
                                finally { try { Marshal.ReleaseComObject(nameIns); } catch { } }
                            }
                            catch { /* défaut 11 */ }

                            int nbInsert = nbMesures - 2;
                            // Insère en bloc : Range(rows pointInsertionRow..pointInsertionRow+nbInsert-1).Insert
                            try
                            {
                                string plage = $"{pointInsertionRow}:{pointInsertionRow + nbInsert - 1}";
                                dynamic rowsRange = nouvelleFeuille.Rows[plage];
                                try { rowsRange.Insert(-4121); }   // xlShiftDown
                                finally { try { Marshal.ReleaseComObject(rowsRange); } catch { } }
                            }
                            catch (Exception ex)
                            {
                                JournalLog.Warn(CategorieLog.Excel, "AJOUT_FEUILLE_INSERTROWS_KO",
                                    $"InsertRowsAbove COM échoué : {ex.Message} — "
                                    + $"fallback sur insertion ligne par ligne.");
                                for (int k = 0; k < nbInsert; k++)
                                {
                                    try
                                    {
                                        dynamic r = nouvelleFeuille.Rows[pointInsertionRow];
                                        try { r.Insert(-4121); }
                                        finally { try { Marshal.ReleaseComObject(r); } catch { } }
                                    }
                                    catch { /* best-effort */ }
                                }
                            }
                        }

                        // --- 11. Formules F (conversion) + G (delta) + N (tachy) pour lignes 11..N ---
                        // Identique à ExcelService.PreparerLignesMesureAsync mais en COM.
                        bool conversionTrMin = estTachy;
                        for (int i = 2; i < nbMesures; i++)
                        {
                            int row = 9 + i;
                            string fC = $"=IF(ISBLANK(ZNCoeffMult),B{row},"
                                + $"(((B{row}-10000000)/(POWER(10,ZNCoeffMult)*10000000))+1)*ZNValFNominale)";
                            EcrireFormuleCelluleInterne(nouvelleFeuille, $"C{row}", fC);
                            EcrireFormuleCelluleInterne(nouvelleFeuille, $"D{row}", $"=C{row - 1}-C{row}");
                            if (conversionTrMin)
                                EcrireFormuleCelluleInterne(nouvelleFeuille, $"K{row}", $"=B{row}*60");
                        }

                        // --- 12. n°Module : plus écrit par ligne (colonne retirée) ; il est
                        //         affiché une seule fois dans la cellule « Module = » (étape 7).

                        // --- 13. Formules métier (ZNFreqMoyReel, ZNVariance, ZNEcartType, …) ---
                        // Posées via Worksheet.Names → RefersToRange.Formula (les Names sheet-scope
                        // créés à l'étape 5 résolvent vers les cellules du template — F13 = moyenne,
                        // F18 = variance, F19 = écart-type, etc. selon le layout ModFeuille).
                        int ligneDeb = 9;
                        int ligneFin = 9 + nbMesures - 1;
                        EcrireFormuleZoneNommeeInterne(nomNouvelleFeuille, "ZNFreqMoyReel",
                            $"=IF(ISBLANK(ZNNbMesures),,AVERAGE(C{ligneDeb}:C{ligneFin}))");

                        // ⚠ PIÈGE EXCEL COM : la formule `=[1]!Cal_variance(SUMSQ(...),...,...)`
                        // est rejetée (0x800A03EC) par .Formula/.Formula2/.FormulaLocal en COM
                        // dynamic, alors qu'elle marche via ClosedXML (écriture XML directe).
                        // On contourne en CALCULANT la variance en C# à la fin de mesure (via
                        // Application.Run pour appeler la macro Cal_variance VBA) et en écrivant
                        // le résultat comme VALEUR dans ZNVariance — cf. FinaliserMesureViaComAsync.
                        //
                        // MAIS : Excel COM rejette aussi la pose des formules `=IF(ISBLANK(ZNVariance),...)`
                        // si ZNVariance pointe vers une cellule TOTALEMENT VIDE — Excel essaie
                        // de résoudre la référence à la pose et plante. Solution : initialiser
                        // ZNVariance avec une valeur 0 placeholder AVANT de poser les formules
                        // qui en dépendent. FinaliserMesureViaComAsync écrasera cette valeur par
                        // la vraie variance, et Excel recalculera ZNEcartType / ZNIncertEcartType
                        // automatiquement.
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNVariance", 0.0);

                        EcrireFormuleZoneNommeeInterne(nomNouvelleFeuille, "ZNEcartType",
                            "=IF(ISBLANK(ZNVariance),,[1]!Cal_ecart_type(ZNVariance))");
                        EcrireFormuleZoneNommeeInterne(nomNouvelleFeuille, "ZNIncertEcartType",
                            "=IF(ISBLANK(ZNNbMesures),,[1]!Cal_incert_ecart_type(ZNVariance,ZNNbMesures))");
                        EcrireFormuleZoneNommeeInterne(nomNouvelleFeuille, "ZNFreqCorr",
                            "=IF(ISBLANK(ZNFreqMoyReel),,[1]!Cal_freq_corrigee(ZNFreqMoyReel,ZNFreqRef))");

                        // --- 14. Forcer format General sur cellules de formules ---
                        // Sans ça, si le format hérité est Texte, Excel garde la formule littéralement
                        // (cf. EcrireFormuleRecapInterne pour le même fix sur Récap.).
                        try
                        {
                            dynamic rangeFormules = nouvelleFeuille.Range[$"C{ligneDeb}:D{ligneFin}"];
                            try { rangeFormules.NumberFormat = "General"; }
                            finally { try { Marshal.ReleaseComObject(rangeFormules); } catch { } }
                        }
                        catch { }

                        // --- 15. Active la feuille (utilisateur la voit immédiatement) ---
                        nouvelleFeuille.Activate();

                        // Re-masque toute shell éventuelle (Activate peut réveiller le window
                        // manager Excel) + assure que le timer périodique est armé.
                        MasquerFenetresShellExcelInterne();
                        DemarrerTimerMasquageShellInterne();

                        // Synchronise le tracking de feuille active dans l'instance
                        if (_feuilleMesure != null)
                        {
                            try { Marshal.ReleaseComObject(_feuilleMesure); } catch { }
                        }
                        _feuilleMesure = nouvelleFeuille;

                        // --- 16. Save (mise à jour fichier disque, Workbook reste ouvert) ---
                        try { _classeurActif.Save(); }
                        catch (Exception ex)
                        {
                            JournalLog.Warn(CategorieLog.Excel, "AJOUT_FEUILLE_SAVE_KO",
                                $"Save COM après ajout feuille échoué : {ex.Message}");
                        }

                        JournalLog.Info(CategorieLog.Excel, "AJOUT_FEUILLE_OK",
                            $"Nouvelle feuille « {nomNouvelleFeuille} » créée via COM "
                            + $"(Workbook reste ouvert — pas de shell SDI).");

                        return nomNouvelleFeuille;
                    }
                    finally
                    {
                        // nouvelleFeuille reste dans _feuilleMesure → on ne release PAS ici.
                        try { if (modFeuille != null) Marshal.ReleaseComObject(modFeuille); } catch { }

                        // Ferme tous les workbooks parasites qui auraient pu être créés malgré
                        // notre paramètre nommé After: (ceinture+bretelles si une version d'Excel
                        // exotique ne respecte pas la convention ; en pratique = no-op si After:
                        // a fonctionné correctement).
                        try { FermerClasseursParasitesInterne(); } catch { }

                        // Restaure les flags Excel : Calculation d'abord (déclenche le recalc
                        // des formules métier qu'on vient de poser) puis EnableEvents puis
                        // ScreenUpdating en dernier (le rendu final apparaît d'un coup, sans flash).
                        try { _excel.Calculation = calculationOriginal; } catch { }
                        try { _excel.EnableEvents = enableEventsOriginal; } catch { }
                        try { _excel.ScreenUpdating = screenUpdatingOriginal; } catch { }

                        // Re-masque toute shell qui aurait pu apparaître malgré le freeze
                        // (sécurité — le timer le ferait dans 500ms max de toute façon).
                        try { MasquerFenetresShellExcelInterne(); } catch { }

                        // CRUCIAL : force la fenêtre du rapport au premier plan. Sans ça, Excel
                        // SDI peut laisser une window grise temporaire en TopMost par-dessus
                        // le rapport (= ce que l'utilisateur voit comme "fenêtre grise qui se
                        // ferme quand on clique sur l'autre Excel manuellement").
                        try { ForcerRapportAuPremierPlanInterne(); } catch { }
                    }
                }
            });

        /// <summary>
        /// Réinitialise le classeur ouvert en supprimant toutes les feuilles de mesure (freq*,
        /// stab*, inter*, topti*, tcont*, strob*, avinter, fqfinale) tout en gardant ModFeuille
        /// + Récap. — l'équivalent d'un "fichier vierge" SANS fermer le Workbook. Évite la shell
        /// SDI qui apparaissait quand l'utilisateur cliquait "Oui = écraser" dans Relancer
        /// (le code historique faisait FermerClasseurActif + File.Delete = ~1-3s à 0 Workbook).
        ///
        /// Tout est encadré par ScreenUpdating=false pour qu'aucun flash visuel ne soit perçu.
        /// Le Récap. est aussi nettoyé (lignes après l'entête) car ses formules cross-sheet
        /// pointant vers les feuilles supprimées deviendraient #REF! sinon.
        /// </summary>
        public Task ReinitialiserClasseurActifAsync()
            => Task.Run(() =>
            {
                lock (_sync)
                {
                    if (_excel == null || _classeurActif == null)
                    {
                        JournalLog.Warn(CategorieLog.Excel, "REINIT_CLASSEUR_SKIP",
                            "ReinitialiserClasseurActifAsync appelé sans classeur actif — ignoré.");
                        return;
                    }

                    // Freeze léger (cf. AjouterFeuilleMesureAsync) — pas de ScreenUpdating=false
                    // qui causerait l'effet "écran gris" pendant la suppression des feuilles.
                    bool enableEventsOriginal = true;
                    int calculationOriginal = -4105;
                    try { enableEventsOriginal = (bool)_excel.EnableEvents; } catch { }
                    try { calculationOriginal = (int)_excel.Calculation; } catch { }
                    try { _excel.EnableEvents = false; } catch { }
                    try { _excel.DisplayAlerts = false; } catch { }
                    try { _excel.Calculation = -4135; } catch { }

                    try
                    {
                        // 1. Énumère les feuilles à supprimer (toutes sauf ModFeuille et Récap.).
                        //    On collecte d'abord les noms avant de supprimer pour éviter de muter
                        //    la collection pendant qu'on itère.
                        var aSupprimer = new List<string>();
                        try
                        {
                            int nb = (int)_classeurActif.Worksheets.Count;
                            for (int i = 1; i <= nb; i++)
                            {
                                try
                                {
                                    dynamic ws = _classeurActif.Worksheets[i];
                                    try
                                    {
                                        string nom = (string)ws.Name;
                                        // Conservation : ModFeuille (modèle de copie) + Récap. (vue
                                        // d'ensemble). On supprime toutes les autres feuilles
                                        // (freq1, stab1, inter1, etc. + anciennes "1".."10" résiduelles).
                                        if (!string.Equals(nom, "ModFeuille", StringComparison.OrdinalIgnoreCase)
                                            && !string.Equals(nom, "Récap.", StringComparison.OrdinalIgnoreCase))
                                        {
                                            aSupprimer.Add(nom);
                                        }
                                    }
                                    finally { try { Marshal.ReleaseComObject(ws); } catch { } }
                                }
                                catch { /* feuille corrompue — skip */ }
                            }
                        }
                        catch (Exception ex)
                        {
                            JournalLog.Warn(CategorieLog.Excel, "REINIT_CLASSEUR_LIST_KO",
                                $"Énumération des feuilles à supprimer échouée : {ex.Message}");
                        }

                        // 2. Suppression effective (DisplayAlerts=false évite la confirmation Excel).
                        int nbSupprimees = 0;
                        foreach (string nom in aSupprimer)
                        {
                            try
                            {
                                dynamic ws = _classeurActif.Worksheets.Item(nom);
                                try { ws.Delete(); nbSupprimees++; }
                                finally { try { Marshal.ReleaseComObject(ws); } catch { } }
                            }
                            catch { /* déjà supprimée ou protection — non bloquant */ }
                        }

                        // 3. Nettoyage Récap. : on supprime les lignes de données (à partir de la
                        //    ligne 6, l'entête étant en ligne 5 pour la Récap. Fréquence ; on
                        //    nettoie largement pour couvrir Stab aussi). Sans ça, les formules
                        //    cross-sheet `='freq1'!ZN*` deviennent #REF! après suppression.
                        try
                        {
                            dynamic recap = _classeurActif.Worksheets.Item("Récap.");
                            try
                            {
                                try { recap.Unprotect("METROL"); }
                                catch { try { recap.Unprotect("metrol"); } catch { try { recap.Unprotect(); } catch { } } }

                                // Efface les lignes 6 à 200 (large mais Excel ignore les vides)
                                try
                                {
                                    dynamic plage = recap.Range["A6:Z200"];
                                    try { plage.Clear(); }
                                    finally { try { Marshal.ReleaseComObject(plage); } catch { } }
                                }
                                catch { }
                            }
                            finally { try { Marshal.ReleaseComObject(recap); } catch { } }
                        }
                        catch { /* Récap. absent : pas de souci */ }

                        // 4. Save : persiste l'état "vierge" sur disque (le fichier existe toujours
                        //    avec son chemin original, mais ne contient plus que ModFeuille + Récap. propre).
                        try { _classeurActif.Save(); }
                        catch (Exception ex)
                        {
                            JournalLog.Warn(CategorieLog.Excel, "REINIT_CLASSEUR_SAVE_KO",
                                $"Save après nettoyage échoué : {ex.Message}");
                        }

                        // _feuilleMesure pointait sur une feuille supprimée → on l'invalide.
                        // La prochaine AjouterFeuilleMesureAsync recréera freq1.
                        if (_feuilleMesure != null)
                        {
                            try { Marshal.ReleaseComObject(_feuilleMesure); } catch { }
                            _feuilleMesure = null;
                        }

                        JournalLog.Info(CategorieLog.Excel, "REINIT_CLASSEUR_OK",
                            $"Classeur réinitialisé via COM : {nbSupprimees} feuille(s) supprimée(s), "
                            + "Récap. nettoyé (Workbook reste ouvert — pas de shell SDI).");
                    }
                    finally
                    {
                        try { _excel.Calculation = calculationOriginal; } catch { }
                        try { _excel.EnableEvents = enableEventsOriginal; } catch { }
                        try { _excel.DisplayAlerts = false; } catch { }
                        try { MasquerFenetresShellExcelInterne(); } catch { }
                        try { ForcerRapportAuPremierPlanInterne(); } catch { }
                    }
                }
            });

        /// <summary>
        /// Génère un nom de feuille unique (freq2, stab3, inter1, …) en parcourant les feuilles
        /// déjà présentes dans le classeur actif. Reproduit <c>ExcelService.TrouverNomFeuilleUnique</c>.
        /// </summary>
        /// <summary>
        /// Supprime la feuille de mesure <paramref name="nomFeuille"/> du classeur COM actif
        /// (et sa ligne Récap.) puis sauvegarde. Appelé quand une mesure est STOPPÉE : la feuille
        /// créée au démarrage ne doit pas être conservée (mesure traitée comme nulle).
        /// Best-effort, ne lève jamais.
        /// </summary>
        public Task SupprimerFeuilleMesureAsync(string nomFeuille) => Task.Run(() =>
        {
            if (string.IsNullOrWhiteSpace(nomFeuille)) return;
            lock (_sync)
            {
                if (_classeurActif == null) return;
                try
                {
                    SupprimerFeuilleEtRecapInterne(nomFeuille);

                    // La feuille courante était celle qu'on vient de supprimer → on lâche la
                    // référence COM pour éviter tout accès ultérieur sur un objet déconnecté.
                    if (_feuilleMesure != null)
                    {
                        try { Marshal.ReleaseComObject(_feuilleMesure); } catch { }
                        _feuilleMesure = null;
                    }

                    try { _classeurActif.Save(); }
                    catch (Exception ex)
                    {
                        JournalLog.Warn(CategorieLog.Excel, "EXCEL_SUPPR_FEUILLE_SAVE_KO",
                            $"Sauvegarde après suppression de « {nomFeuille} » échouée : {ex.Message}");
                    }

                    JournalLog.Info(CategorieLog.Excel, "EXCEL_FEUILLE_STOP_SUPPRIMEE",
                        $"Feuille « {nomFeuille} » supprimée suite à l'arrêt de la mesure.");
                }
                catch (Exception ex)
                {
                    JournalLog.Warn(CategorieLog.Excel, "EXCEL_SUPPR_FEUILLE_STOP_KO",
                        $"Suppression de la feuille « {nomFeuille} » (arrêt) échouée : {ex.Message}");
                }
            }
        });

        /// <summary>
        /// Supprime du classeur COM actif la feuille <paramref name="nomFeuille"/> si elle existe,
        /// ainsi que les lignes de la Récap. qui la référencent. À appeler SOUS LOCK (_sync).
        /// Gère DisplayAlerts pour éviter la confirmation Excel de suppression de feuille. No-op si
        /// la feuille n'existe pas (cas des feuilles numérotées au nom neuf).
        /// </summary>
        private void SupprimerFeuilleEtRecapInterne(string nomFeuille)
        {
            if (_classeurActif == null || string.IsNullOrWhiteSpace(nomFeuille)) return;

            // La feuille existe-t-elle ? Sinon no-op.
            bool existe = false;
            try
            {
                dynamic wsTest = _classeurActif.Worksheets.Item(nomFeuille);
                existe = wsTest != null;
                try { Marshal.ReleaseComObject(wsTest); } catch { }
            }
            catch { existe = false; }
            if (!existe) return;

            bool alertesOriginal = true;
            try { alertesOriginal = (bool)_excel.DisplayAlerts; } catch { }
            try { _excel.DisplayAlerts = false; } catch { }
            try
            {
                // 1. Retire les lignes de la Récap. qui référencent la feuille (sinon doublon ou
                //    #REF! après suppression / recréation d'une feuille de même nom).
                try
                {
                    dynamic recap = _classeurActif.Worksheets.Item("Récap.");
                    try
                    {
                        try { recap.Unprotect("METROL"); }
                        catch { try { recap.Unprotect("metrol"); } catch { try { recap.Unprotect(); } catch { } } }

                        // Borne le balayage à la zone utilisée (évite 100×12 lectures COM inutiles).
                        int maxLigne = 60;
                        try
                        {
                            dynamic used = recap.UsedRange;
                            int firstRow = (int)used.Row;
                            int rowCount = (int)used.Rows.Count;
                            maxLigne = firstRow + rowCount - 1;
                            try { Marshal.ReleaseComObject(used); } catch { }
                        }
                        catch { maxLigne = 60; }
                        if (maxLigne > 200) maxLigne = 200;

                        // Descendant : supprimer une ligne ne décale pas les lignes < r non visitées.
                        for (int r = maxLigne; r >= 6; r--)
                        {
                            if (!LigneRecapRefereFeuilleInterne(recap, r, nomFeuille)) continue;
                            try
                            {
                                dynamic cellA = recap.Range[$"A{r}"];
                                try
                                {
                                    dynamic entireRow = cellA.EntireRow;
                                    try { entireRow.Delete(); }
                                    finally { try { Marshal.ReleaseComObject(entireRow); } catch { } }
                                }
                                finally { try { Marshal.ReleaseComObject(cellA); } catch { } }
                            }
                            catch { /* ligne inaccessible — skip */ }
                        }
                    }
                    finally { try { Marshal.ReleaseComObject(recap); } catch { } }
                }
                catch { /* pas de Récap. — non bloquant */ }

                // 2. Supprime la feuille elle-même.
                try
                {
                    dynamic ws = _classeurActif.Worksheets.Item(nomFeuille);
                    try { ws.Delete(); }
                    finally { try { Marshal.ReleaseComObject(ws); } catch { } }
                }
                catch (Exception ex)
                {
                    JournalLog.Warn(CategorieLog.Excel, "EXCEL_SUPPR_FEUILLE_KO",
                        $"Suppression de la feuille « {nomFeuille} » échouée : {ex.Message}");
                }
            }
            finally
            {
                try { _excel.DisplayAlerts = alertesOriginal; } catch { }
            }
        }

        /// <summary>Vrai si une cellule (col 1..12) de la ligne Récap. contient une formule référençant la feuille.</summary>
        private bool LigneRecapRefereFeuilleInterne(dynamic recap, int row, string nomFeuille)
        {
            for (int c = 1; c <= 12; c++)
            {
                try
                {
                    dynamic cell = recap.Cells[row, c];
                    try
                    {
                        string f = "";
                        try { f = (string)cell.Formula; } catch { }
                        if (!string.IsNullOrEmpty(f)
                            && (f.IndexOf($"{nomFeuille}!", StringComparison.OrdinalIgnoreCase) >= 0
                                || f.IndexOf($"'{nomFeuille}'!", StringComparison.OrdinalIgnoreCase) >= 0))
                            return true;
                    }
                    finally { try { Marshal.ReleaseComObject(cell); } catch { } }
                }
                catch { }
            }
            return false;
        }

        private string TrouverNomFeuilleUniqueInterne(TypeMesure type)
        {
            if (_classeurActif == null) return "Mesure1";
            if (type == TypeMesure.FreqAvantInterv) return "avinter";
            if (type == TypeMesure.FreqFinale) return "fqfinale";

            string prefixe = type switch
            {
                TypeMesure.Frequence    => "freq",
                TypeMesure.Stabilite    => "stab",
                TypeMesure.Interval     => "inter",
                TypeMesure.TachyOptique => "topti",
                TypeMesure.TachyContact => "tcont",
                TypeMesure.Stroboscope  => "strob",
                _ => "Mesure"
            };

            // Collecte des noms existants
            var existants = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                int nb = (int)_classeurActif.Worksheets.Count;
                for (int i = 1; i <= nb; i++)
                {
                    try
                    {
                        dynamic ws = _classeurActif.Worksheets[i];
                        try { existants.Add((string)ws.Name); }
                        finally { try { Marshal.ReleaseComObject(ws); } catch { } }
                    }
                    catch { }
                }
            }
            catch { }

            int idx = 1;
            while (existants.Contains($"{prefixe}{idx}")) idx++;
            return $"{prefixe}{idx}";
        }

        /// <summary>
        /// Écrit une valeur dans une cellule via <c>Range["A1"].Value2</c>. Utilitaire interne
        /// pour les écritures en-têtes/labels de <see cref="AjouterFeuilleMesureAsync"/>.
        /// </summary>
        private static void EcrireValeurCelluleInterne(dynamic feuille, string adresseA1, object valeur)
        {
            try
            {
                dynamic c = feuille.Range[adresseA1];
                try { c.Value2 = valeur; }
                finally { try { Marshal.ReleaseComObject(c); } catch { } }
            }
            catch { /* best-effort */ }
        }

        /// <summary>Écrit une formule dans une cellule via <c>Range["A1"].Formula</c>.</summary>
        private static void EcrireFormuleCelluleInterne(dynamic feuille, string adresseA1, string formule)
        {
            try
            {
                dynamic c = feuille.Range[adresseA1];
                try
                {
                    try { c.NumberFormat = "General"; } catch { }
                    c.Formula = formule;
                }
                finally { try { Marshal.ReleaseComObject(c); } catch { } }
            }
            catch { /* best-effort */ }
        }

        /// <summary>
        /// Écrit une formule dans la cellule pointée par une zone nommée sheet-scope.
        /// Équivalent de <see cref="EcrireValeurZoneNommeeInterne"/> mais pour
        /// <c>RefersToRange.Formula</c>. Si l'écriture échoue (typiquement parce que la formule
        /// commence par <c>[1]!</c> et que le link externe n'est pas encore résolu à ce stade,
        /// ou parce que la formule contient une syntaxe que <c>.Formula</c> rejette mais que
        /// <c>.Formula2</c>/<c>.FormulaLocal</c> accepte), on essaie plusieurs variantes en
        /// cascade et on logge l'erreur précise — pour ne plus jamais avoir un échec silencieux
        /// comme celui de <c>ZNVariance</c> qui restait vide.
        /// </summary>
        private void EcrireFormuleZoneNommeeInterne(string nomFeuille, string nomZone, string formule)
        {
            if (_classeurActif == null) return;
            string echecsAccumules = "";
            try
            {
                dynamic ws = _classeurActif.Worksheets.Item(nomFeuille);
                try
                {
                    dynamic? nom = null;
                    try { nom = ws.Names.Item(nomZone); }
                    catch (Exception exNom)
                    {
                        echecsAccumules += $"Names.Item KO: {exNom.Message}; ";
                    }

                    bool succes = false;
                    if (nom != null)
                    {
                        try
                        {
                            dynamic refRange = nom.RefersToRange;
                            try
                            {
                                try { refRange.NumberFormat = "General"; } catch { }
                                // Essai 1 : .Formula (US, virgules) — méthode standard.
                                try
                                {
                                    refRange.Formula = formule;
                                    succes = true;
                                }
                                catch (Exception ex1)
                                {
                                    echecsAccumules += $".Formula KO: {ex1.Message}; ";
                                    // Essai 2 : .Formula2 (Excel 365+, supporte arrays dynamiques).
                                    try
                                    {
                                        refRange.Formula2 = formule;
                                        succes = true;
                                    }
                                    catch (Exception ex2)
                                    {
                                        echecsAccumules += $".Formula2 KO: {ex2.Message}; ";
                                        // Essai 3 : .FormulaLocal (séparateurs FR : point-virgule).
                                        try
                                        {
                                            string formuleLocal = formule.Replace(",", ";");
                                            refRange.FormulaLocal = formuleLocal;
                                            succes = true;
                                            JournalLog.Info(CategorieLog.Excel, "FORMULE_VIA_FORMULALOCAL",
                                                $"Pose formule {nomZone} via FormulaLocal (séparateurs FR) "
                                                + $"après échec Formula/Formula2.");
                                        }
                                        catch (Exception ex3)
                                        {
                                            echecsAccumules += $".FormulaLocal KO: {ex3.Message}; ";
                                        }
                                    }
                                }
                            }
                            finally { try { Marshal.ReleaseComObject(refRange); } catch { } }
                        }
                        finally { try { Marshal.ReleaseComObject(nom); } catch { } }
                    }

                    if (!succes)
                    {
                        // Fallback workbook-scope
                        try
                        {
                            dynamic nomWb = _classeurActif.Names.Item(nomZone);
                            try
                            {
                                dynamic refRangeWb = nomWb.RefersToRange;
                                try
                                {
                                    try { refRangeWb.NumberFormat = "General"; } catch { }
                                    try { refRangeWb.Formula = formule; succes = true; }
                                    catch (Exception exWb)
                                    {
                                        echecsAccumules += $"WB.Formula KO: {exWb.Message}; ";
                                    }
                                }
                                finally { try { Marshal.ReleaseComObject(refRangeWb); } catch { } }
                            }
                            finally { try { Marshal.ReleaseComObject(nomWb); } catch { } }
                        }
                        catch (Exception exWbName) { echecsAccumules += $"WB.Names KO: {exWbName.Message}; "; }
                    }

                    if (!succes)
                    {
                        JournalLog.Warn(CategorieLog.Excel, "FORMULE_POSE_KO",
                            $"Échec pose formule '{nomZone}' = \"{formule}\" sur '{nomFeuille}'. "
                            + $"Tentatives : {echecsAccumules}");
                    }
                }
                finally { try { Marshal.ReleaseComObject(ws); } catch { } }
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "FORMULE_POSE_ERREUR_GLOBALE",
                    $"Pose formule '{nomZone}' = \"{formule}\" sur '{nomFeuille}' : "
                    + $"erreur globale {ex.GetType().Name}: {ex.Message}");
            }
        }

        /// <summary>Reproduit <c>ExcelService.NomAppareilDepuisCatalogue</c> en local.</summary>
        private static string NomAppareilDepuisCatalogueLocal(string? idModele)
        {
            if (string.IsNullOrEmpty(idModele)) return string.Empty;
            var modele = Metrologo.Services.Catalogue.CatalogueAppareilsService.Instance.Modeles
                .FirstOrDefault(m => m.Id == idModele);
            return modele?.Nom ?? idModele;
        }

        /// <summary>
        /// Logue le scope effectif (sheet vs workbook) d'un sous-ensemble critique de zones
        /// nommées sur la nouvelle feuille. Détecte les cas où le clonage Names sheet-scope
        /// aurait échoué silencieusement (les formules métier seraient alors résolues vers
        /// ModFeuille = données vides, comme dans la précédente tentative option A).
        /// </summary>
        private void VerifierNamesSheetScopeInterne(string nomFeuille)
        {
            if (_classeurActif == null) return;

            // Sous-ensemble critique = celles dont la résolution sheet-scope est obligatoire pour
            // que la mesure soit calculée correctement.
            string[] zonesCritiques =
            {
                "ZNFreqMoyReel", "ZNVariance", "ZNEcartType", "ZNIncertEcartType", "ZNFreqCorr",
                "ZNNoFiche", "ZNDate", "ZNTypeMesure", "ZNNbMesures", "ZNFreqRef",
                "ZNCoeffA", "ZNCoeffB", "ZNValFNominale", "ZNCoeffMult"
            };

            var manquants = new List<string>();
            try
            {
                dynamic ws = _classeurActif.Worksheets.Item(nomFeuille);
                try
                {
                    foreach (string zone in zonesCritiques)
                    {
                        try
                        {
                            dynamic n = ws.Names.Item(zone);
                            try { /* OK : sheet-scope présent */ }
                            finally { try { Marshal.ReleaseComObject(n); } catch { } }
                        }
                        catch { manquants.Add(zone); }
                    }
                }
                finally { try { Marshal.ReleaseComObject(ws); } catch { } }
            }
            catch { /* abandon validation */ }

            if (manquants.Count > 0)
            {
                JournalLog.Warn(CategorieLog.Excel, "AJOUT_FEUILLE_NAMES_MANQUANTS",
                    $"Zones nommées sheet-scope manquantes sur {nomFeuille} : "
                    + string.Join(", ", manquants)
                    + " — les formules métier qui y font référence risquent d'être KO.");
            }
            else
            {
                JournalLog.Info(CategorieLog.Excel, "AJOUT_FEUILLE_NAMES_VALIDES",
                    $"Toutes les zones nommées critiques ({zonesCritiques.Length}) "
                    + $"sont bien sheet-scope sur {nomFeuille}.");
            }
        }

        /// <summary>
        /// Calcule la variance en appelant la macro VBA <c>Cal_variance</c> du Metrologo.xla
        /// via <c>Application.Run</c>, puis écrit le résultat comme valeur littérale dans
        /// <c>ZNVariance</c>. Workaround au rejet Excel 0x800A03EC de la formule
        /// <c>=[1]!Cal_variance(SUMSQ(...),...,...)</c> via <c>.Formula</c> en COM dynamic.
        ///
        /// Reproduit en C# le calcul que ferait la formule :
        ///   - <c>SUMSQ(G{deb+1}:G{fin})</c> = somme des carrés des deltas de fréquence réelle
        ///   - On convertit chaque mesure brute en fréquence réelle (mode Direct/Indirect) pour
        ///     reproduire la colonne F du rapport
        ///   - On passe ces 3 arguments à Cal_variance qui retourne la variance Allan-like
        ///
        /// Si Application.Run échoue (xla pas chargé, macro absente, etc.), on fallback sur
        /// un calcul C# pur <c>sumSq / (2*(N-1))</c> = formule Allan classique.
        /// </summary>
        private void CalculerEtEcrireVarianceViaComInterne(Mesure mesure, List<double> resultats,
            string nomFeuille)
        {
            if (_excel == null || _classeurActif == null) return;
            if (resultats == null || resultats.Count < 2) return;

            try
            {
                // 1. Reproduit la colonne F (fréq. réelle) en C#, comme le ferait la formule
                //    `=IF(ISBLANK(ZNCoeffMult),E,(((E-1e7)/(POWER(10,Mult)*1e7))+1)*FNominale)`.
                var corriges = new List<double>(resultats.Count);
                foreach (var v in resultats) corriges.Add(ConvertirEnFreqReelleLocal(v, mesure));

                // 2. SUMSQ des deltas (G_{i+1} = F_i - F_{i+1}, donc on somme les carrés des
                //    différences consécutives à partir de l'indice 1).
                double sumSq = 0;
                for (int i = 1; i < corriges.Count; i++)
                {
                    double delta = corriges[i - 1] - corriges[i];
                    sumSq += delta * delta;
                }

                double moyenne = 0;
                foreach (var v in corriges) moyenne += v;
                moyenne /= corriges.Count;
                int nbMesures = corriges.Count;

                // 3. Appel macro VBA Cal_variance via Application.Run.
                //    Syntaxe COM : Application.Run("Metrologo.xla!Cal_variance", arg1, arg2, arg3)
                //    Si pas accessible (Run non disponible, macro absente), fallback C# pur.
                double? variance = null;
                try
                {
                    dynamic res = _excel.Run("Metrologo.xla!Cal_variance",
                        sumSq, (double)nbMesures, moyenne);
                    variance = Convert.ToDouble(res);
                    JournalLog.Info(CategorieLog.Excel, "VARIANCE_VIA_RUN_OK",
                        $"Cal_variance({sumSq:G6}, {nbMesures}, {moyenne:G6}) = {variance:G6} "
                        + $"(via Application.Run)");
                }
                catch (Exception exRun)
                {
                    // Fallback : calcul C# pur Allan variance σ²(τ) = sumSq / (2 * (N-1))
                    if (nbMesures >= 2)
                    {
                        variance = sumSq / (2.0 * (nbMesures - 1));
                        JournalLog.Warn(CategorieLog.Excel, "VARIANCE_VIA_RUN_FALLBACK",
                            $"Application.Run KO ({exRun.Message}) — fallback C# Allan : "
                            + $"variance = sumSq/(2*(N-1)) = {variance:G6}");
                    }
                }

                // 4. Écrit la valeur numérique dans ZNVariance.
                if (variance.HasValue)
                {
                    EcrireValeurZoneNommeeInterne(nomFeuille, "ZNVariance", variance.Value);
                }
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "VARIANCE_CALCUL_KO",
                    $"Calcul variance échoué : {ex.Message}");
            }
        }

        /// <summary>
        /// Calcule les coefficients d'incertitude depuis le module sélectionné et les écrit
        /// dans les zones nommées <c>ZNCoeffA/B</c> (ou <c>C/D</c> pour tachymètres) via COM.
        /// Porte la logique de <c>ExcelService.EcrireStatsAsync</c> mais en écriture COM directe.
        /// </summary>
        private void EcrireCoeffsIncertitudeViaComInterne(Mesure mesure, List<double> resultats,
            string nomFeuille, double tempsGateSecondes)
        {
            // Estimation sigma relatif (n-1 classique) pour Stabilité : calibre l'axe Y log
            // du graphe Stab à la dernière gate. Identique à la logique de
            // ExcelService.EcrireStatsAsync — équivalent en C# pur, pas d'écriture Excel.
            if (mesure.TypeMesure == TypeMesure.Stabilite && resultats.Count >= 2)
            {
                int i = nomFeuille.Length;
                while (i > 0 && char.IsDigit(nomFeuille[i - 1])) i--;
                string chiffres = nomFeuille.Substring(i);
                if (int.TryParse(chiffres, out int numGate))
                {
                    double moy = resultats.Average();
                    double sumSq = 0;
                    foreach (var v in resultats) sumSq += (v - moy) * (v - moy);
                    double sigma = Math.Sqrt(sumSq / (resultats.Count - 1));
                    if (moy > 0 && sigma > 0)
                    {
                        _sigmasRelatifsParGate[numGate] = sigma / moy;
                    }
                }
            }

            if (string.IsNullOrEmpty(mesure.NumModuleIncertitude) || resultats.Count == 0) return;

            double moyenneBrute = resultats.Average();
            double moyenneReelle = ConvertirEnFreqReelleLocal(moyenneBrute, mesure);
            bool uniteRpm = EnTetesMesureHelper.EstUniteRpm(mesure.TypeMesure);
            double valeurLookup = uniteRpm ? moyenneReelle * 60.0 : moyenneReelle;

            string fonction = IncertitudeFonctionHelper.NomFonction(mesure.TypeMesure);
            var (coeffA, coeffB) = ModulesIncertitudeService.ObtenirCoefficients(
                mesure.NumModuleIncertitude, mesure.TypeMesure, fonction,
                tempsGateSecondes, valeurLookup);

            if (coeffA > 0 || coeffB > 0)
            {
                if (EnTetesMesureHelper.EstTachymetre(mesure.TypeMesure))
                {
                    EcrireValeurZoneNommeeInterne(nomFeuille, "ZNCoeffC", coeffA);
                    EcrireValeurZoneNommeeInterne(nomFeuille, "ZNCoeffD", coeffB);
                }
                else
                {
                    EcrireValeurZoneNommeeInterne(nomFeuille, "ZNCoeffA", coeffA);
                    EcrireValeurZoneNommeeInterne(nomFeuille, "ZNCoeffB", coeffB);
                }
                JournalLog.Info(CategorieLog.Mesure, "INCERT_COEFFS_RESOLUS_COM",
                    $"Module {mesure.NumModuleIncertitude} : CoeffA={coeffA:G6} CoeffB={coeffB:G6} "
                    + $"(fonction={fonction}, gate={tempsGateSecondes}s, valeurLookup={valeurLookup:G6})");
            }

            // Module Fréquence auxiliaire (tachymètres uniquement)
            if (EnTetesMesureHelper.EstTachymetre(mesure.TypeMesure)
                && !string.IsNullOrEmpty(mesure.NumModuleIncertitudeFreq))
            {
                string fonctionFreq = IncertitudeFonctionHelper.NomFonction(TypeMesure.Frequence);
                var (coeffAFreq, coeffBFreq) = ModulesIncertitudeService.ObtenirCoefficients(
                    mesure.NumModuleIncertitudeFreq, TypeMesure.Frequence, fonctionFreq,
                    tempsGateSecondes, moyenneReelle);
                if (coeffAFreq > 0 || coeffBFreq > 0)
                {
                    EcrireValeurZoneNommeeInterne(nomFeuille, "ZNCoeffA", coeffAFreq);
                    EcrireValeurZoneNommeeInterne(nomFeuille, "ZNCoeffB", coeffBFreq);
                }
            }
        }

        /// <summary>Reproduit ExcelService.ConvertirEnFreqReelle en local pour le COM path.</summary>
        private static double ConvertirEnFreqReelleLocal(double brute, Mesure mesure)
        {
            if (mesure.ModeMesure == ModeMesure.Direct) return brute;
            return (((brute - 10_000_000.0)
                    / (Math.Pow(10, mesure.IndexMultiplicateur) * 10_000_000.0)) + 1.0)
                   * mesure.FNominale;
        }

        /// <summary>
        /// Redirige tous les liens externes pointant vers un fichier <c>*.xla</c> du classeur
        /// actif vers <see cref="Preferences.CheminMacroXLA"/> via la méthode COM
        /// <c>Workbook.ChangeLink</c>. Excel nettoie alors proprement le rels (1 seul
        /// Relationship par lien) et résout immédiatement les formules <c>[1]!Cal_xxx(...)</c>.
        /// Remplace le PatcherLienMacroXLA (ClosedXML) qui n'est plus appelé en mode visible.
        /// </summary>
        private void RedirigerLinksXlaViaComInterne()
        {
            if (_excel == null || _classeurActif == null) return;
            string xlaPath = Preferences.CheminMacroXLA;
            if (string.IsNullOrWhiteSpace(xlaPath)) return;
            try
            {
                // LinkSources(xlLinkTypeExcelLinks=1) retourne array 1-based des chemins liés.
                dynamic? links = null;
                try { links = _classeurActif.LinkSources(1); } catch { /* aucun lien */ }
                if (links == null) return;

                int nb = links.Length;
                for (int i = 1; i <= nb; i++)
                {
                    try
                    {
                        string ancien = (string)links.GetValue(i);
                        if (ancien.IndexOf("Metrologo.xla", StringComparison.OrdinalIgnoreCase) < 0)
                            continue;
                        if (string.Equals(ancien, xlaPath, StringComparison.OrdinalIgnoreCase))
                            continue;   // déjà au bon chemin

                        _classeurActif.ChangeLink(ancien, xlaPath, 1);   // 1 = xlLinkTypeExcelLinks
                        JournalLog.Info(CategorieLog.Excel, "XLA_LINK_REDIRIGE",
                            $"Lien externe redirigé via COM : {ancien} → {xlaPath}");
                    }
                    catch (Exception ex)
                    {
                        JournalLog.Warn(CategorieLog.Excel, "XLA_LINK_REDIRIGE_KO",
                            $"ChangeLink échoué : {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "XLA_LINK_REDIRIGE_GLOBAL_KO",
                    $"Patch lien xla via COM échoué : {ex.Message}");
            }
        }

        /// <summary>
        /// Écrit une valeur dans une zone nommée workbook-scope du classeur actif via COM.
        /// Utilisé pour les zones partagées entre feuilles (ex: ZNNoFicheRecapFreq qui couvre
        /// l'en-tête de la feuille Récap).
        /// </summary>
        private void EcrireValeurZoneNommeeClasseurInterne(string nomZone, object valeur)
        {
            if (_classeurActif == null) return;
            try
            {
                dynamic nom = _classeurActif.Names.Item(nomZone);
                try { nom.RefersToRange.Value2 = valeur; }
                finally { try { Marshal.ReleaseComObject(nom); } catch { } }
            }
            catch { /* zone absente — silencieux */ }
        }

        /// <summary>Écrit une valeur dans une zone nommée sheet-scope via COM.</summary>
        private void EcrireValeurZoneNommeeInterne(string nomFeuille, string nomZone, object valeur)
        {
            if (_classeurActif == null) return;
            try
            {
                dynamic ws = _classeurActif.Worksheets.Item(nomFeuille);
                try
                {
                    dynamic nom = ws.Names.Item(nomZone);
                    try
                    {
                        nom.RefersToRange.Value2 = valeur;
                    }
                    finally { try { Marshal.ReleaseComObject(nom); } catch { } }
                }
                catch
                {
                    // Fallback workbook-scope si la zone n'est pas sur la feuille
                    try
                    {
                        dynamic nomWb = _classeurActif.Names.Item(nomZone);
                        try { nomWb.RefersToRange.Value2 = valeur; }
                        finally { try { Marshal.ReleaseComObject(nomWb); } catch { } }
                    }
                    catch { /* zone absente, silencieux */ }
                }
                finally { try { Marshal.ReleaseComObject(ws); } catch { } }
            }
            catch { /* best-effort */ }
        }

        /// <summary>Mot de passe prédéfini de protection des feuilles (verrouillage post-mesure).</summary>
        private const string MotDePasseProtectionFeuille = "METROL";

        /// <summary>
        /// Re-protège TOUTES les feuilles du classeur actif (feuilles de mesure + Récap.) avec
        /// <see cref="MotDePasseProtectionFeuille"/>. Appelé en fin de mesure, juste avant le Save :
        /// le rapport n'est plus modifiable dans Excel sans le mot de passe.
        ///
        /// L'app, qui connaît ce mot de passe, déprotège automatiquement à chaque relance / ajout
        /// de mesure (AjouterFeuilleMesureAsync déprotège la nouvelle feuille, les mises à jour
        /// Récap. déprotègent la feuille Récap.). Chaque feuille est d'abord déprotégée (idempotent —
        /// Protect lève si déjà protégée) puis re-protégée. Best-effort : un échec sur une feuille
        /// n'interrompt ni les autres ni la sauvegarde.
        /// </summary>
        private void ProtegerToutesFeuillesActifInterne()
        {
            if (_classeurActif == null) return;

            dynamic? sheets = null;
            try
            {
                sheets = _classeurActif.Worksheets;
                int count = (int)sheets.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic? ws = null;
                    try
                    {
                        ws = sheets.Item(i);
                        try { ws.Unprotect(MotDePasseProtectionFeuille); }
                        catch { try { ws.Unprotect("metrol"); } catch { try { ws.Unprotect(); } catch { } } }
                        ws.Protect(MotDePasseProtectionFeuille);
                    }
                    catch (Exception ex)
                    {
                        JournalLog.Warn(CategorieLog.Excel, "PROTECTION_FEUILLE_KO",
                            $"Protection d'une feuille via COM échouée : {ex.Message}");
                    }
                    finally { if (ws != null) { try { Marshal.ReleaseComObject(ws); } catch { } } }
                }
                JournalLog.Info(CategorieLog.Excel, "FEUILLES_PROTEGEES",
                    $"{count} feuille(s) re-protégée(s) (mot de passe prédéfini) en fin de mesure.");
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "PROTECTION_CLASSEUR_KO",
                    $"Protection des feuilles échouée : {ex.Message}");
            }
            finally { if (sheets != null) { try { Marshal.ReleaseComObject(sheets); } catch { } } }
        }

        /// <summary>Ajoute la ligne Récap Fréquence via COM (équivalent MettreAJourRecapFreqAsync).</summary>
        private void AjouterLigneRecapFreqViaComInterne(Mesure mesure, string nomFeuille)
        {
            if (_classeurActif == null) return;
            dynamic? recap = null;
            try { recap = _classeurActif.Worksheets.Item("Récap."); } catch { return; }
            if (recap == null) return;

            try
            {
                // Déprotéger
                try { recap.Unprotect("METROL"); } catch { try { recap.Unprotect("metrol"); } catch { try { recap.Unprotect(); } catch { } } }

                // Élargit la colonne D (date) pour qu'elle s'affiche en entier au lieu de « #### ».
                try { recap.Columns["D"].ColumnWidth = 13.0; } catch { }

                // En-tête fiche : N° FI + date du jour. Écraser les valeurs héritées du
                // template (sinon le récap affiche un FI/date qui n'a rien à voir avec la
                // mesure en cours).
                EcrireValeurZoneNommeeClasseurInterne("ZNNoFicheRecapFreq", mesure.NumFI);
                EcrireValeurZoneNommeeClasseurInterne("ZNDateRecapFreq",
                    DateTime.Now.ToString("dd/MM/yyyy"));

                // Trouver la 1ère ligne vide après l'entête (ligne 5)
                int ligneEntete = 5;
                int nouvelleLigne = ligneEntete + 1;
                while (true)
                {
                    bool vide = true;
                    for (int c = 1; c <= 12; c++)
                    {
                        try
                        {
                            dynamic cell = recap.Cells[nouvelleLigne, c];
                            object? val = cell.Value2;
                            Marshal.ReleaseComObject(cell);
                            if (val != null && val.ToString() != "") { vide = false; break; }
                        }
                        catch { }
                    }
                    if (vide || nouvelleLigne > 100) break;
                    nouvelleLigne++;
                }

                string qf = $"'{nomFeuille}'!";
                // Col 1 : ZNFreqMoyReel | 2 : ZNLibGate | 3 : ZNFreqCorr | 4 : ZNEcartType
                EcrireFormuleRecapInterne(recap, nouvelleLigne, 1,  $"={qf}ZNFreqMoyReel");
                EcrireFormuleRecapInterne(recap, nouvelleLigne, 2,  $"={qf}ZNLibGate");
                EcrireFormuleRecapInterne(recap, nouvelleLigne, 3,  $"={qf}ZNFreqCorr");
                EcrireFormuleRecapInterne(recap, nouvelleLigne, 4,  $"={qf}ZNEcartType");
                // Col 5 : fréquence indiquée. Générateur = « Géné. » (écrit direct).
                // Fréquencemètre = valeur saisie en post-mesure, écrite DIRECTEMENT dans
                // cette cellule par EcrireFreqIndiqueeRecapAsync (la pop-up) → on laisse
                // la cellule vide ici, plus aucune formule vers ZNFreqRef de la feuille.
                if (mesure.SourceMesure == SourceMesure.Generateur)
                    EcrireValeurRecapInterne(recap, nouvelleLigne, 5,  "Géné.");
                // Col 6 : ZNIncertResol | 7 : ZNIncertSup | 8 : ZNIncertAccreditee | 9 : ZNIncertGlobale
                EcrireFormuleRecapInterne(recap, nouvelleLigne, 6,  $"={qf}ZNIncertResol");
                EcrireFormuleRecapInterne(recap, nouvelleLigne, 7,  $"={qf}ZNIncertSup");
                EcrireFormuleRecapInterne(recap, nouvelleLigne, 8,  $"={qf}ZNIncertAccreditee");
                EcrireFormuleRecapInterne(recap, nouvelleLigne, 9,  $"={qf}ZNIncertGlobale");
                // Col 10 : fréquence finale calculée
                int n = nouvelleLigne;
                EcrireFormuleRecapInterne(recap, nouvelleLigne, 10, $"=IF(ISNUMBER(E{n}),(E{n}/C{n})*A{n},C{n})");
                // Col 11 : n°Module — lu depuis la cellule « Module = » de la feuille (F10, ou
                // G10 en tachy). La colonne Fonction a été retirée.
                string celModuleRecap = EnTetesMesureHelper.EstTachymetre(mesure.TypeMesure) ? "G10" : "F10";
                EcrireFormuleRecapInterne(recap, nouvelleLigne, 11, $"={qf}{celModuleRecap}");

                // Écritures scientifiques (écart-type + incertitudes) : forcées à 2 décimales.
                FormaterScientifiqueRecapInterne(recap, nouvelleLigne, 4, 6, 7, 8, 9);

                // En-tête K si vide
                EcrireValeurRecapSiVideInterne(recap, ligneEntete, 11, "n°Module");
            }
            finally
            {
                try { Marshal.ReleaseComObject(recap); } catch { }
            }
        }

        /// <summary>
        /// Écrit la « fréquence indiquée » saisie en post-mesure (mode Fréquencemètre)
        /// DIRECTEMENT dans la colonne 5 de la ligne Récap de la feuille de mesure active —
        /// à l'emplacement exact où le mode Générateur écrit « Géné. ».
        ///
        /// Ne touche PAS la feuille de mesure : la valeur n'est qu'un report d'affichage
        /// dans le récap, elle n'alimente plus ZNFreqRef ni la fréquence corrigée (ZNFreqCorr).
        /// La bonne ligne est identifiée par la formule de sa colonne 1
        /// (<c>='&lt;feuille&gt;'!ZNFreqMoyReel</c>) qui référence la feuille active.
        /// </summary>
        public Task EcrireFreqIndiqueeRecapAsync(double valeur) => Task.Run(() =>
        {
            lock (_sync)
            {
                if (_classeurActif == null || _feuilleMesure == null) return;

                string nomFeuille;
                try { nomFeuille = (string)_feuilleMesure.Name; }
                catch { return; }

                dynamic? recap = null;
                try { recap = _classeurActif.Worksheets.Item("Récap."); } catch { return; }
                if (recap == null) return;

                try
                {
                    try { recap.Unprotect("METROL"); }
                    catch { try { recap.Unprotect("metrol"); } catch { try { recap.Unprotect(); } catch { } } }

                    // Parcourt les lignes de données (à partir de la ligne 6) et repère
                    // celle dont la col 1 référence la feuille active.
                    for (int row = 6; row <= 100; row++)
                    {
                        dynamic? cell = null;
                        try
                        {
                            cell = recap.Cells[row, 1];
                            string formule = "";
                            try { formule = (string)(cell.Formula ?? ""); } catch { }

                            if (formule.IndexOf("'" + nomFeuille + "'!", StringComparison.OrdinalIgnoreCase) >= 0
                                || formule.IndexOf(nomFeuille + "!", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                EcrireValeurRecapInterne(recap, row, 5, valeur);
                                JournalLog.Info(CategorieLog.Excel, "RECAP_FREQ_INDIQUEE",
                                    $"Fréquence indiquée {valeur} écrite en col 5 du récap (ligne {row}, feuille {nomFeuille}).");
                                break;
                            }
                        }
                        catch { }
                        finally { if (cell != null) { try { Marshal.ReleaseComObject(cell); } catch { } } }
                    }
                }
                finally
                {
                    try { Marshal.ReleaseComObject(recap); } catch { }
                }
            }
        });

        /// <summary>Ajoute la ligne Récap Stabilité via COM (équivalent MettreAJourRecapStabAsync).</summary>
        private void AjouterLigneRecapStabViaComInterne(Mesure mesure, string nomFeuille)
        {
            if (_classeurActif == null) return;
            dynamic? recap = null;
            try { recap = _classeurActif.Worksheets.Item("Récap."); } catch { return; }
            if (recap == null) return;

            try
            {
                try { recap.Unprotect("METROL"); } catch { try { recap.Unprotect("metrol"); } catch { try { recap.Unprotect(); } catch { } } }

                // Élargit la colonne D (date) pour qu'elle s'affiche en entier au lieu de « #### ».
                try { recap.Columns["D"].ColumnWidth = 13.0; } catch { }

                // En-tête fiche
                EcrireValeurRecapInterne(recap, 1, 2, mesure.NumFI);
                EcrireValeurRecapInterne(recap, 1, 4, DateTime.Now.ToString("dd/MM/yyyy"));

                // Numéro de gate (parsé depuis le nom de feuille — fallback compteur)
                int numero;
                int i = nomFeuille.Length;
                while (i > 0 && char.IsDigit(nomFeuille[i - 1])) i--;
                string chiffres = nomFeuille.Substring(i);
                if (!int.TryParse(chiffres, out numero))
                {
                    // Compte les lignes remplies (col A non vide)
                    int compteur = 0;
                    for (int r = 6; r <= 18; r++)
                    {
                        try
                        {
                            dynamic cell = recap.Cells[r, 1];
                            object? val = cell.Value2;
                            Marshal.ReleaseComObject(cell);
                            if (val != null && val.ToString() != "") compteur++;
                        }
                        catch { }
                    }
                    numero = compteur + 1;
                }

                int ligne = 5 + numero;
                string qf = $"'{nomFeuille}'!";
                EcrireFormuleRecapInterne(recap, ligne, 1, $"={qf}ZNValGateSecondes");
                EcrireFormuleRecapInterne(recap, ligne, 2, $"={qf}ZNFreqMoyReel");
                EcrireFormuleRecapInterne(recap, ligne, 3, $"={qf}ZNEcartType");
                EcrireFormuleRecapInterne(recap, ligne, 4, $"={qf}ZNIncertEcartType");
                EcrireFormuleRecapInterne(recap, ligne, 5, $"={qf}ZNIncertAccreditee");
                EcrireFormuleRecapInterne(recap, ligne, 6, $"={qf}ZNIncertGlobale");
                EcrireFormuleRecapInterne(recap, ligne, 7, $"=C{ligne}+F{ligne}");
                EcrireFormuleRecapInterne(recap, ligne, 8,
                    $"=IF(ISBLANK(C{ligne}),,IF((C{ligne}-F{ligne})<=0,0,C{ligne}-F{ligne}))");
                // n°Module lu depuis la cellule « Module = » (F10) ; colonne Fonction retirée.
                EcrireFormuleRecapInterne(recap, ligne, 11, $"={qf}F10");

                // Écritures scientifiques (écart-type + incertitudes) : forcées à 2 décimales.
                FormaterScientifiqueRecapInterne(recap, ligne, 3, 4, 5, 6);

                EcrireValeurRecapSiVideInterne(recap, 5, 11, "n°Module");
            }
            finally
            {
                try { Marshal.ReleaseComObject(recap); } catch { }
            }
        }

        /// <summary>
        /// Force le format scientifique à 2 décimales (<c>0.00E+00</c>) sur les cellules indiquées
        /// d'une ligne du récap (colonnes d'incertitude / écart-type). Best-effort par cellule.
        /// </summary>
        private static void FormaterScientifiqueRecapInterne(dynamic recap, int row, params int[] colonnes)
        {
            foreach (int col in colonnes)
            {
                try
                {
                    dynamic c = recap.Cells[row, col];
                    c.NumberFormat = "0.00E+00";
                    Marshal.ReleaseComObject(c);
                }
                catch { /* best-effort */ }
            }
        }

        private static void EcrireFormuleRecapInterne(dynamic recap, int row, int col, string formule)
        {
            try
            {
                dynamic c = recap.Cells[row, col];
                try
                {
                    // CAUSE RACINE du bug "ligne 1 du Récap. en texte" :
                    // Le template METROLOGO.xlsx a une ligne 6 pré-stylée avec des cellules
                    // marquées t="s" (shared string) dans le XML xlsx. Quand la 1ère mesure
                    // tombe sur cette ligne, l'écriture .Formula = "=..." STOCKE LA FORMULE
                    // COMME STRING (entrée dans sharedStrings.xml) au lieu d'une vraie formule.
                    // Résultat : Excel affiche « ='freq1'!ZNLibGate » littéralement.
                    //
                    // NumberFormat = "General" et Clear() ne changent pas le type stocké (t="s"
                    // reste). La SEULE façon de forcer le type numérique est d'écrire un nombre
                    // dans la cellule AVANT de poser la formule. Ceci convertit t="s" en t="n"
                    // (numeric), et la pose de formule suivante stocke correctement.
                    try { c.Clear(); } catch { }
                    try { c.Style = "Normal"; } catch { }
                    try { c.NumberFormat = "General"; } catch { }
                    try { c.Value2 = 0.0; } catch { }   // CRUCIAL : force type numeric (t="s" → t="n")
                    c.Formula = formule;

                    // Centrage horizontal + vertical pour cohérence visuelle avec le reste
                    // du Récap (Clear() a effacé l'alignement hérité du template).
                    // xlCenter = -4108 (constante Excel COM).
                    try { c.HorizontalAlignment = -4108; } catch { }
                    try { c.VerticalAlignment = -4108; } catch { }

                    // Vérif post-pose : ceinture+bretelles. Si malgré tout la cellule est
                    // restée en string, on bypass via Application.Evaluate (perte du recalc
                    // live mais affichage correct).
                    try
                    {
                        object v = c.Value2;
                        if (v is string s && s.StartsWith("=", StringComparison.Ordinal))
                        {
                            try
                            {
                                dynamic app = c.Application;
                                object resultat = app.Evaluate(formule.Substring(1));
                                if (resultat != null)
                                {
                                    c.NumberFormat = "General";
                                    c.Value2 = resultat;
                                    JournalLog.Info(CategorieLog.Excel, "RECAP_CELL_FALLBACK_VALUE",
                                        $"Cell [{row},{col}] : fallback Evaluate (formule restée en texte).");
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
                finally { try { Marshal.ReleaseComObject(c); } catch { } }
            }
            catch { }
        }

        private static void EcrireValeurRecapInterne(dynamic recap, int row, int col, object valeur)
        {
            try { dynamic c = recap.Cells[row, col]; c.Value2 = valeur; Marshal.ReleaseComObject(c); }
            catch { }
        }

        private static void EcrireValeurRecapSiVideInterne(dynamic recap, int row, int col, object valeur)
        {
            try
            {
                dynamic c = recap.Cells[row, col];
                object? val = c.Value2;
                if (val == null || val.ToString() == "") c.Value2 = valeur;
                Marshal.ReleaseComObject(c);
            }
            catch { }
        }

        /// <summary>
        /// Ajuste les plages des séries du graphe Stab au nombre réel de gates balayées —
        /// équivalent COM de <c>ExcelService.RendreGrapheStabDynamique</c>. Patche XValues
        /// et Values via la Chart Object Model d'Excel (pas de manipulation XML zippée
        /// puisque le fichier reste ouvert).
        /// </summary>
        private void AjusterPlagesGrapheStabViaComInterne(int nbGates)
        {
            if (nbGates <= 0 || _classeurActif == null) return;
            int derniereLigne = 5 + nbGates;
            try
            {
                dynamic ws = _classeurActif.Worksheets.Item("Récap.");
                try
                {
                    dynamic chartObjects = ws.ChartObjects();
                    int nbCharts = chartObjects.Count;
                    for (int i = 1; i <= nbCharts; i++)
                    {
                        dynamic chartObj = chartObjects.Item(i);
                        try
                        {
                            dynamic chart = chartObj.Chart;
                            dynamic seriesColl = chart.SeriesCollection();
                            int nbSeries = seriesColl.Count;
                            for (int j = 1; j <= nbSeries; j++)
                            {
                                dynamic serie = seriesColl.Item(j);
                                try
                                {
                                    // Patch des plages via regex sur les chaînes de référence.
                                    string xRef = (string)serie.Formula;
                                    string nouvelle = System.Text.RegularExpressions.Regex.Replace(
                                        xRef,
                                        @"\$([A-Z]+)\$6:\$([A-Z]+)\$\d+",
                                        m => $"${m.Groups[1].Value}$6:${m.Groups[2].Value}${derniereLigne}");
                                    if (nouvelle != xRef) serie.Formula = nouvelle;
                                }
                                catch { /* série non patchable */ }
                                finally { try { Marshal.ReleaseComObject(serie); } catch { } }
                            }
                            // Axe Y auto-scale (retire les bornes figées 1E-12/1E-08 du Stab1.xls)
                            try
                            {
                                dynamic valAxis = chart.Axes(2);   // 2 = xlValue
                                valAxis.MinimumScaleIsAuto = true;
                                valAxis.MaximumScaleIsAuto = true;
                                try { Marshal.ReleaseComObject(valAxis); } catch { }
                            }
                            catch { }
                            try { Marshal.ReleaseComObject(seriesColl); } catch { }
                        }
                        finally { try { Marshal.ReleaseComObject(chartObj); } catch { } }
                    }
                    try { Marshal.ReleaseComObject(chartObjects); } catch { }
                }
                finally { try { Marshal.ReleaseComObject(ws); } catch { } }
                JournalLog.Info(CategorieLog.Excel, "GRAPHE_STAB_AJUSTE_COM",
                    $"Plages séries graphe Stab ajustées via COM (nbGates={nbGates}, dernière ligne=L{derniereLigne}).");
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "GRAPHE_STAB_AJUSTE_KO",
                    $"Ajustement plages graphe Stab via COM échoué : {ex.Message}");
            }
        }

        /// <summary>
        /// Calibre l'axe Y log du graphe Stab à partir des sigmas relatifs accumulés —
        /// équivalent COM de <c>ExcelService.CalibrerAxeYGrapheStab</c>. Définit
        /// <c>MinimumScale</c> et <c>MaximumScale</c> via la Chart Object Model.
        /// </summary>
        private void CalibrerAxeYGrapheStabViaComInterne(IEnumerable<double> sigmas)
        {
            if (_classeurActif == null) return;
            var positifs = sigmas.Where(s => s > 0).ToList();
            if (positifs.Count == 0) return;

            double minS = positifs.Min();
            double maxS = positifs.Max();
            double minBorne = Math.Pow(10, Math.Floor(Math.Log10(minS) - 0.5));
            double maxBorne = Math.Pow(10, Math.Ceiling(Math.Log10(maxS) + 0.5));
            if (maxBorne / minBorne < 10) maxBorne = minBorne * 10;

            try
            {
                dynamic ws = _classeurActif.Worksheets.Item("Récap.");
                try
                {
                    dynamic chartObjects = ws.ChartObjects();
                    int nbCharts = chartObjects.Count;
                    for (int i = 1; i <= nbCharts; i++)
                    {
                        dynamic chartObj = chartObjects.Item(i);
                        try
                        {
                            dynamic chart = chartObj.Chart;
                            try
                            {
                                dynamic valAxis = chart.Axes(2);   // xlValue
                                valAxis.MinimumScale = minBorne;
                                valAxis.MaximumScale = maxBorne;
                                try { Marshal.ReleaseComObject(valAxis); } catch { }
                            }
                            catch (Exception ex)
                            {
                                JournalLog.Warn(CategorieLog.Excel, "GRAPHE_STAB_AXE_Y_KO",
                                    $"Calibration axe Y graphe Stab échouée : {ex.Message}");
                            }
                        }
                        finally { try { Marshal.ReleaseComObject(chartObj); } catch { } }
                    }
                    try { Marshal.ReleaseComObject(chartObjects); } catch { }
                }
                finally { try { Marshal.ReleaseComObject(ws); } catch { } }
                JournalLog.Info(CategorieLog.Excel, "GRAPHE_STAB_AXE_Y_OK",
                    $"Axe Y graphe Stab calibré via COM : [{minBorne:E2} ; {maxBorne:E2}].");
            }
            catch { /* best-effort */ }
        }

        /// <summary>
        /// Ferme proprement notre instance Excel COM (qui contient le <c>.xla</c> pré-chargé
        /// causant la fenêtre Frame « grise ») et ouvre ensuite le rapport <paramref name="cheminFichier"/>
        /// via le shell Windows (équivalent d'un double-clic utilisateur). Excel se charge
        /// alors avec 1 SEULE fenêtre propre, fusionnée Application+Workbook, sans pollution.
        /// La prochaine mesure relancera automatiquement une nouvelle instance COM via
        /// <see cref="EstInstanceVivante"/> + <see cref="RedemarrerInstanceInterne"/>.
        /// </summary>
        public Task QuitterEtOuvrirViaShellAsync(string cheminFichier) => Task.Run(() =>
        {
            // OPTION C v2 : on NE quitte PAS notre instance Excel COM (qui détient le .xla
            // pré-chargé). On ferme juste le Workbook actif et on laisse l'instance vivre
            // en background avec le .xla en mémoire. Process.Start envoie le rapport à Excel ;
            // si Excel détecte notre instance déjà en cours, il y ouvre le rapport dedans
            // → le .xla est trouvé → [1]!Cal_xxx résolus, pas de popup « Liaisons ».
            lock (_sync)
            {
                try
                {
                    FermerClasseurActifInterne();
                    // L'instance Excel et le .xla restent en mémoire pour servir le shell open.
                }
                catch (Exception ex)
                {
                    JournalLog.Warn(CategorieLog.Excel, "EXCEL_FERME_CLASSEUR_KO",
                        $"Fermeture du classeur actif échouée : {ex.Message}");
                }
            }
            // Petit délai pour laisser Excel libérer le handle fichier avant le shell open.
            System.Threading.Thread.Sleep(300);

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = cheminFichier,
                    UseShellExecute = true
                });
                JournalLog.Info(CategorieLog.Excel, "EXCEL_OUVERT_VIA_SHELL",
                    $"Rapport ouvert via shell (instance Excel + .xla en mémoire) : {Path.GetFileName(cheminFichier)}");
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "EXCEL_OUVERT_VIA_SHELL_KO",
                    $"Ouverture via shell échouée : {ex.Message}");
            }
        });

        /// <summary>
        /// Déplace les fenêtres « shell » Excel hors écran (-32000,-32000) AVANT que
        /// <c>Visible=true</c> ne les rende visibles. Élimine le clignotement résiduel
        /// du <c>ShowWindow(SW_HIDE)</c> post-Visible (où la shell est brièvement visible
        /// avant d'être masquée). Combiné ensuite avec SW_HIDE comme ceinture+bretelles.
        /// </summary>
        private void DeplacerHorsEcranFenetresShellInterne()
        {
            if (_excel == null || _excelPid <= 0 || _classeurActif == null) return;
            try
            {
                IntPtr hwndRapport = IntPtr.Zero;
                try
                {
                    dynamic w = _classeurActif.Windows[1];
                    try { hwndRapport = new IntPtr((int)w.Hwnd); }
                    finally { try { Marshal.ReleaseComObject(w); } catch { } }
                }
                catch { /* best-effort */ }

                IntPtr hwndRapportLocal = hwndRapport;
                EnumWindows((hWnd, _) =>
                {
                    try
                    {
                        GetWindowThreadProcessId(hWnd, out uint pid);
                        if (pid != (uint)_excelPid) return true;
                        if (hWnd == hwndRapportLocal) return true;
                        if (GetParent(hWnd) != IntPtr.Zero) return true;
                        // Déplacement hors écran — Excel ne repositionne pas la fenêtre quand
                        // il fait son ShowWindow(SW_SHOW) interne, donc elle reste invisible.
                        SetWindowPos(hWnd, IntPtr.Zero, -32000, -32000, 0, 0,
                            SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
                    }
                    catch { /* best-effort */ }
                    return true;
                }, IntPtr.Zero);
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "EXCEL_SHELL_OFFSCREEN_KO",
                    $"Déplacement hors écran des fenêtres shell échoué : {ex.Message}");
            }
        }

        /// <summary>
        /// Énumère TOUTES les fenêtres top-level du process Excel et logge HWND, titre, visible,
        /// IsIconic, position, taille. Utilisé pour diagnostiquer précisément quelle fenêtre
        /// est la "grise" perçue par l'utilisateur — si on voit dans le log une window avec
        /// titre "Excel" (sans nom de fichier) qui reste visible et au premier plan, c'est elle.
        /// Logge même les fenêtres invisibles pour repérer celles qui se réveillent.
        /// </summary>
        private void DiagnostiquerFenetresExcelInterne(string moment)
        {
            if (_excelPid <= 0) return;
            try
            {
                var lignes = new List<string>();
                IntPtr foreground = GetForegroundWindow();

                EnumWindows((hWnd, _) =>
                {
                    try
                    {
                        GetWindowThreadProcessId(hWnd, out uint pid);
                        if (pid != (uint)_excelPid) return true;

                        // Récupère titre, visibilité, parent, rect.
                        var sb = new System.Text.StringBuilder(256);
                        GetWindowText(hWnd, sb, 256);
                        string titre = sb.ToString();
                        bool visible = IsWindowVisible(hWnd);
                        bool topLevel = GetParent(hWnd) == IntPtr.Zero;
                        bool isForeground = (hWnd == foreground);
                        string rect = "??";
                        if (GetWindowRect(hWnd, out RECT r))
                            rect = $"{r.Left},{r.Top}->{r.Right - r.Left}x{r.Bottom - r.Top}";

                        lignes.Add($"  hwnd=0x{hWnd.ToInt64():X} "
                            + $"visible={(visible ? "Y" : "N")} "
                            + $"top={(topLevel ? "Y" : "N")} "
                            + $"fg={(isForeground ? "Y" : "N")} "
                            + $"rect=[{rect}] "
                            + $"titre=\"{titre}\"");
                    }
                    catch { }
                    return true;
                }, IntPtr.Zero);

                JournalLog.Info(CategorieLog.Excel, "DIAG_FENETRES_EXCEL",
                    $"[{moment}] {lignes.Count} fenêtre(s) Excel PID={_excelPid} :\n"
                    + string.Join("\n", lignes));
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "DIAG_FENETRES_KO",
                    $"Diag fenêtres échoué : {ex.Message}");
            }
        }

        /// <summary>
        /// Force la fenêtre du rapport actif au premier plan visuel (SetForegroundWindow +
        /// BringWindowToTop + WindowState=xlMaximized) — annule l'effet "fenêtre grise par-dessus"
        /// que produit Excel SDI quand certaines opérations COM (Copy, Save, Open) mettent
        /// momentanément une shell Excel en TopMost.
        ///
        /// Le piège classique de SetForegroundWindow sous Windows : il échoue silencieusement
        /// si le thread appelant n'a pas le focus. La technique standard pour contourner :
        /// attacher temporairement le thread courant au thread foreground via AttachThreadInput,
        /// faire l'opération, puis détacher.
        /// </summary>
        private void ForcerRapportAuPremierPlanInterne()
        {
            if (_excel == null || _classeurActif == null) return;
            try
            {
                // 1. Activate côté Excel COM : amène la fenêtre du rapport au top du Z-order
                //    Excel (mais ne touche pas le Z-order Windows).
                try { _classeurActif.Activate(); } catch { }

                // 2. WindowState = xlMaximized : la fenêtre du rapport prend tout l'écran Excel,
                //    masquant visuellement toute autre window Excel en dessous.
                try
                {
                    dynamic w = _classeurActif.Windows[1];
                    try { w.WindowState = -4137; /* xlMaximized */ }
                    finally { try { Marshal.ReleaseComObject(w); } catch { } }
                }
                catch { }

                // 3. SetForegroundWindow + BringWindowToTop côté Win32 : amène la fenêtre du
                //    rapport au top du Z-order Windows, devant TOUTE autre window (y compris
                //    la shell Excel TopMost qui apparaissait par-dessus). Trick AttachThreadInput
                //    pour éviter l'échec silencieux de SetForegroundWindow quand notre process
                //    n'a pas le focus à ce moment-là.
                IntPtr hwndRapport = IntPtr.Zero;
                try
                {
                    dynamic w = _classeurActif.Windows[1];
                    try { hwndRapport = new IntPtr((int)w.Hwnd); }
                    finally { try { Marshal.ReleaseComObject(w); } catch { } }
                }
                catch { }
                if (hwndRapport == IntPtr.Zero) return;

                try
                {
                    IntPtr hwndForeground = GetForegroundWindow();
                    uint pidIgnore;
                    uint threadForeground = GetWindowThreadProcessId(hwndForeground, out pidIgnore);
                    uint threadCourant = GetCurrentThreadId();
                    bool attached = false;
                    try
                    {
                        if (threadForeground != threadCourant)
                        {
                            attached = AttachThreadInput(threadCourant, threadForeground, true);
                        }
                        BringWindowToTop(hwndRapport);
                        SetForegroundWindow(hwndRapport);
                    }
                    finally
                    {
                        if (attached)
                        {
                            try { AttachThreadInput(threadCourant, threadForeground, false); } catch { }
                        }
                    }
                }
                catch { /* best-effort */ }
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "FORCER_PREMIER_PLAN_KO",
                    $"Forcer le rapport au premier plan a échoué : {ex.Message}");
            }
        }

        /// <summary>
        /// Démarre le timer de masquage continu de la shell SDI. Tant qu'un classeur est ouvert
        /// dans Excel, on appelle <see cref="MasquerFenetresShellExcelInterne"/> toutes les 500ms.
        /// Garantit qu'aucune shell ne peut rester visible plus d'une demi-seconde, même si
        /// l'appel initial post-ouverture rate (race COM, ouverture asynchrone du window manager
        /// Excel, etc.). Coût quasi nul : <see cref="MasquerFenetresShellExcelInterne"/> fait
        /// uniquement un EnumWindows Win32 (~quelques µs), sans aucun appel COM.
        ///
        /// ⚠ <see cref="FermerClasseursParasitesInterne"/> N'EST PLUS appelé ici. Cet appel
        ///   itérait sur tous les Workbooks via COM (~100-300 ms par tick) et prenait le lock
        ///   <c>_sync</c> partagé avec <see cref="EcrireValeurLiveAsync"/>. Pendant une mesure,
        ///   chaque tick bloquait l'écriture live de la valeur en cours d'arrivée du compteur
        ///   GPIB → effet « rien ne s'affiche pendant plusieurs secondes puis ça reprend »
        ///   particulièrement visible quand le compteur met du temps à répondre (signal carré,
        ///   retries internes, etc.). Le nettoyage des Workbooks parasites est déjà couvert
        ///   par les appels existants dans <c>OuvrirEtAfficherAsync</c> et
        ///   <c>AjouterFeuilleMesureAsync</c> — les seuls moments où une shell peut apparaître.
        /// </summary>
        private void DemarrerTimerMasquageShellInterne()
        {
            lock (_timerLock)
            {
                if (_timerMasquageShell != null) return;
                _timerMasquageShell = new System.Threading.Timer(_ =>
                {
                    try
                    {
                        lock (_sync)
                        {
                            if (_excel == null || _classeurActif == null) return;
                            // Uniquement MasquerFenetresShellExcelInterne : fait 1 mini call COM
                            // (Hwnd du Workbook actif) + EnumWindows Win32. Tick ~1-2 ms total,
                            // négligeable pour les écritures live concurrentes.
                            MasquerFenetresShellExcelInterne();
                        }
                    }
                    catch { /* best-effort, ne doit jamais lever depuis un Timer */ }
                }, null, 500, 500);
            }
        }

        /// <summary>
        /// Arrête le timer de masquage. À appeler à la fermeture du classeur (plus de risque de
        /// shell parasite) et au Dispose de l'app. Idempotent.
        /// </summary>
        private void ArreterTimerMasquageShellInterne()
        {
            lock (_timerLock)
            {
                _timerMasquageShell?.Dispose();
                _timerMasquageShell = null;
            }
        }

        /// <summary>
        /// Masque la fenêtre « shell » top-level d'Excel (titre = « Excel » sans nom de
        /// fichier) qui apparaît grise à côté du rapport en SDI. Cette fenêtre n'est pas
        /// énumérée par <c>Application.Windows</c> (qui ne liste que les Workbook Windows),
        /// d'où le besoin de descendre au niveau Win32 EnumWindows.
        /// Ne masque QUE les fenêtres dont le titre est « Excel » ou « Microsoft Excel »
        /// pour ne pas affecter le rapport ouvert (titre = « Mesures_F6_xxxxx - Excel »).
        /// </summary>
        private void MasquerFenetresShellExcelInterne()
        {
            if (_excel == null || _excelPid <= 0 || _classeurActif == null) return;
            try
            {
                // Récupère le HWND de la fenêtre du rapport actif — la SEULE fenêtre top-level
                // Excel du process qu'on garde visible. Tout le reste (shell SDI vide, Book1
                // résiduel, etc.) est masqué, peu importe le caption (qui varie selon la
                // langue/version d'Excel : « Excel », « Microsoft Excel », « Book1 », …).
                IntPtr hwndRapport = IntPtr.Zero;
                try
                {
                    dynamic w = _classeurActif.Windows[1];
                    try { hwndRapport = new IntPtr((int)w.Hwnd); }
                    finally { try { Marshal.ReleaseComObject(w); } catch { } }
                }
                catch { /* best-effort, on saura via le log de diag */ }

                // GARDE-FOU CRITIQUE : si on n'a pas pu identifier le HWND du rapport
                // (race COM : Window pas encore créée), on ne masque RIEN. Sinon on
                // masquerait TOUTES les fenêtres Excel y compris le rapport lui-même
                // → l'utilisateur ne voit plus rien et Excel finit par se fermer.
                if (hwndRapport == IntPtr.Zero)
                {
                    JournalLog.Warn(CategorieLog.Excel, "EXCEL_SHELL_MASQUE_SKIP",
                        "HWND du rapport indisponible (race COM) — masquage shell sauté pour éviter de cacher le rapport.");
                    return;
                }

                int nbMasquees = 0;
                IntPtr hwndRapportLocal = hwndRapport;
                EnumWindows((hWnd, _) =>
                {
                    try
                    {
                        GetWindowThreadProcessId(hWnd, out uint pid);
                        if (pid != (uint)_excelPid) return true;        // pas notre process Excel
                        if (!IsWindowVisible(hWnd)) return true;        // déjà masqué
                        if (hWnd == hwndRapportLocal) return true;       // c'est le rapport, on garde
                        if (GetParent(hWnd) != IntPtr.Zero) return true; // pas top-level (enfant)

                        // CEINTURE + BRETELLES + HARNAIS pour s'assurer que la shell SDI
                        // disparaît visuellement, même si elle se ré-affiche derrière notre dos
                        // (Excel SDI peut ré-afficher la shell sur certaines opérations COM —
                        // SaveAs, Workbooks.Open, etc.).
                        // 1. SW_HIDE — la cache complètement.
                        // 2. SW_FORCEMINIMIZE — au cas où elle se ré-afficherait, elle sera
                        //    en barre des tâches minimisée plutôt qu'à l'écran.
                        // 3. SetWindowPos HWND_BOTTOM — la pousse tout en bas du Z-order
                        //    pour que toute fenêtre visible (y compris le rapport) la cache.
                        // 4. Déplacement hors écran (-32000, -32000) — sécurité ultime visuel.
                        try { ShowWindow(hWnd, SW_HIDE); } catch { }
                        try { ShowWindow(hWnd, SW_FORCEMINIMIZE); } catch { }
                        try
                        {
                            SetWindowPos(hWnd, HWND_BOTTOM, 0, 0, 0, 0,
                                SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);
                        }
                        catch { }
                        try
                        {
                            SetWindowPos(hWnd, IntPtr.Zero, -32000, -32000, 0, 0,
                                SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
                        }
                        catch { }
                        nbMasquees++;
                    }
                    catch { /* best-effort */ }
                    return true;
                }, IntPtr.Zero);
                if (nbMasquees > 0)
                {
                    JournalLog.Info(CategorieLog.Excel, "EXCEL_SHELL_MASQUE",
                        $"{nbMasquees} fenêtre(s) parasite(s) Excel masquée(s) — "
                        + $"hwndRapport=0x{hwndRapport.ToInt64():X} "
                        + "(SW_HIDE + SW_FORCEMINIMIZE + HWND_BOTTOM + off-screen).");
                }
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "EXCEL_SHELL_MASQUE_KO",
                    $"Masquage des fenêtres shell Excel échoué : {ex.Message}");
            }
        }

        /// <summary>
        /// Comparaison robuste de chemins : normalise via Path.GetFullPath pour gérer
        /// les variations (capitalisation, slash, raccourcis 8.3) qui faisaient échouer
        /// un simple string.Equals et fermaient à tort le rapport actif.
        /// </summary>
        private static bool ChemPathEgal(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return false;
            try
            {
                string na = Path.GetFullPath(a);
                string nb = Path.GetFullPath(b);
                return string.Equals(na, nb, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
            }
        }

        private void FermerClasseurActifInterne()
        {
            // Arrête le timer de masquage shell : plus de classeur ouvert = plus besoin
            // de scanner les fenêtres parasites.
            ArreterTimerMasquageShellInterne();

            // CRUCIAL pour éviter la shell SDI grise pendant la transition entre 2 mesures :
            // on cache TOUTE l'instance Excel AVANT de fermer le Workbook. Sans ça, Excel
            // resterait Visible=true avec 0 Workbook → shell apparaîtrait. Avec Visible=false,
            // l'instance est totalement invisible Windows pendant que ClosedXML écrit le fichier.
            // OuvrirEtAfficherAsync remettra Visible=true APRÈS avoir ouvert le nouveau Workbook
            // (séquence identique à la mesure 1 qui marche parfaitement).
            if (_excel != null)
            {
                try { _excel.Visible = false; } catch { }
                try { _excel.ScreenUpdating = false; } catch { }
            }

            if (_feuilleMesure != null)
            {
                try { Marshal.ReleaseComObject(_feuilleMesure); } catch { }
                _feuilleMesure = null;
            }

            if (_classeurActif != null)
            {
                try { _classeurActif.Close(true); } catch { }   // SaveChanges = true
                try { Marshal.ReleaseComObject(_classeurActif); } catch { }
                _classeurActif = null;
            }

            _cheminClasseurActif = string.Empty;

            // Force la libération immédiate des handles fichier détenus par Excel via les
            // proxys COM RCW. Sans ces 2 lignes, le fichier reste "verrouillé" côté OS pour
            // 100-500 ms après le Close → ClosedXML bascule sur un nom horodaté à la mesure
            // suivante au lieu d'écrire dans le même fichier.
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // On NE fait plus _excel.Visible=false ici : ça créait un toggle false→true
            // à la mesure suivante (= fenêtre grise SDI au moment du Visible=true).
            // L'instance reste visible avec sa fenêtre principale entre deux mesures ; les
            // fenêtres parasites (shell SDI vide) sont masquées via MasquerFenetresShellExcelInterne
            // au moment de l'ouverture du nouveau Workbook. Le Visible=false n'est plus fait
            // que dans Dispose (à la fermeture définitive de l'app).
        }

        /// <summary>
        /// Quitte l'instance Excel hôte. À appeler à la fermeture de l'application. Les éventuels
        /// classeurs ouverts par l'utilisateur dans cette instance seront fermés (Excel demandera
        /// si besoin de sauver si des modifs sont en cours).
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            lock (_sync)
            {
                FermerClasseurActifInterne();

                if (_excel != null)
                {
                    try { _excel.Quit(); } catch { }
                    try { Marshal.ReleaseComObject(_excel); } catch { }
                    _excel = null;
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
