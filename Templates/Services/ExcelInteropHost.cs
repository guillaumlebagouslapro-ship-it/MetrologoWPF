using Metrologo.Models;
using Metrologo.Services.Incertitude;
using Metrologo.Services.Journal;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services
{
    /// <summary>
    /// Hôte Excel partagé : on garde une seule instance Excel.Application, cachée et lancée au
    /// démarrage, que toutes les mesures réutilisent pour l'affichage live.
    /// </summary>
    // On fait du COM en late-binding (dynamic + GetTypeFromProgID) plutôt que de passer par la PIA
    // Microsoft.Office.Interop : office.dll n'est pas toujours installée selon la version d'Office,
    // et son absence fait planter l'assembly loader ("Could not load file or assembly 'office'").
    // Pense bien à relâcher chaque objet COM avec Marshal.ReleaseComObject, faute de quoi EXCEL.EXE
    // survit en zombie après la fermeture de Metrologo.
    public sealed class ExcelInteropHost : IDisposable
    {
        private static readonly Lazy<ExcelInteropHost> _instance = new(() => new ExcelInteropHost());
        public static ExcelInteropHost Instance => _instance.Value;

        private dynamic? _excel;
        private dynamic? _classeurActif;
        private dynamic? _feuilleMesure;
        private string _cheminClasseurActif = string.Empty;

        /// <summary>Nom complet du classeur actif TEL QU'EXCEL L'ÉCRIT (lettre de lecteur ou
        /// \\serveur\partage selon sa configuration). Sert à reconnaître à coup sûr notre rapport
        /// parmi les classeurs ouverts, quel que soit le dossier réglé dans Admin.</summary>
        private string _fullNameActifExcel = string.Empty;

        /// <summary>
        /// Chemin du Metrologo.xla qu'on pré-ouvre dans Excel : c'est lui qui résout les formules
        /// [1]!Cal_xxx des rapports. On le mémorise pour le mettre à l'abri du nettoyage des classeurs
        /// parasites. Reste vide tant qu'il n'a pas été pré-chargé.
        /// </summary>
        private string _cheminXlaPreCharge = string.Empty;

        /// <summary>
        /// Sigma relatif (n-1) par gate, qu'on accumule au fil d'une session Stab : il sert à
        /// calibrer l'axe Y log du graphe à la dernière gate. On le remet à zéro au début de
        /// chaque nouvelle session.
        /// </summary>
        private readonly Dictionary<int, double> _sigmasRelatifsParGate = new();

        public void ReinitialiserSigmasStab() => _sigmasRelatifsParGate.Clear();

        /// <summary>
        /// Liste les EXCEL.EXE en cours d'exécution, en excluant notre instance cachée (_excelPid).
        /// Ça nous permet de repérer les Excel ouverts par l'utilisateur qui verrouillent le .xlsx
        /// de mesure et font échouer ClosedXML.
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
                    catch { /* le process est mort entre le Get et notre test */ }
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
        /// Termine les process EXCEL.EXE qu'on lui passe. Best-effort, renvoie le nombre fermés.
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

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved,
            out System.Runtime.InteropServices.ComTypes.IRunningObjectTable prot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved,
            out System.Runtime.InteropServices.ComTypes.IBindCtx ppbc);

        /// <summary>
        /// Parcourt TOUTES les instances Excel ouvertes et ferme uniquement les classeurs de MESURE
        /// (nom <c>freq</c>/<c>stab</c> sous <c>…\Metrologo\</c>), qu'on repère PAR LEUR NOM via la
        /// Running Object Table, sans tuer le moindre process Excel. Les autres fichiers ouverts par
        /// l'utilisateur restent intacts, même non sauvegardés. On épargne aussi le classeur de mesure
        /// actif de notre propre instance hôte. Best-effort, renvoie le nombre de classeurs fermés.
        /// </summary>
        public int FermerClasseursMesureOuvertsExternes()
        {
            int fermes = 0;
            System.Runtime.InteropServices.ComTypes.IRunningObjectTable? rot = null;
            System.Runtime.InteropServices.ComTypes.IEnumMoniker? enumMon = null;
            try
            {
                if (GetRunningObjectTable(0, out rot) != 0 || rot == null) return 0;
                rot.EnumRunning(out enumMon);
                enumMon.Reset();

                var monikers = new System.Runtime.InteropServices.ComTypes.IMoniker[1];
                while (enumMon.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    var moniker = monikers[0];
                    try
                    {
                        string? nomAffiche = null;
                        try
                        {
                            CreateBindCtx(0, out var ctx);
                            moniker.GetDisplayName(ctx, null, out nomAffiche);
                            Marshal.ReleaseComObject(ctx);
                        }
                        catch { /* moniker non nommable : on l'ignore */ }

                        // Dans la ROT, le nom affiché d'un classeur Excel correspond à son chemin
                        // complet. On ne ferme QUE les fichiers de mesure (freq/stab sous \Metrologo\),
                        // et surtout jamais le classeur actif de notre instance hôte.
                        if (string.IsNullOrEmpty(nomAffiche)
                            || !EstClasseurDeMesure(nomAffiche)
                            || EstClasseurActif(nomAffiche))
                            continue;

                        object? comObj = null;
                        try { if (rot.GetObject(moniker, out comObj) != 0) comObj = null; }
                        catch { comObj = null; }
                        if (comObj == null) continue;

                        dynamic wb = comObj;
                        try
                        {
                            wb.Close(false);   // on ne sauvegarde pas : la mesure régénère le rapport
                            fermes++;
                            JournalLog.Info(CategorieLog.Excel, "EXCEL_MESURE_FERME_PARNOM",
                                $"Classeur de mesure « {nomAffiche} » fermé (par nom) avant la mesure.");
                        }
                        catch (Exception exClose)
                        {
                            JournalLog.Warn(CategorieLog.Excel, "EXCEL_MESURE_FERME_PARNOM_KO",
                                $"Fermeture du classeur « {nomAffiche} » échouée : {exClose.Message}");
                        }
                        finally { try { Marshal.ReleaseComObject(wb); } catch { } }
                    }
                    finally { try { Marshal.ReleaseComObject(moniker); } catch { } }
                }
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "EXCEL_ROT_KO",
                    $"Parcours de la Running Object Table échoué : {ex.Message}");
            }
            finally
            {
                if (enumMon != null) { try { Marshal.ReleaseComObject(enumMon); } catch { } }
                if (rot != null) { try { Marshal.ReleaseComObject(rot); } catch { } }
            }
            return fermes;
        }

        /// <summary>
        /// Photo des PID EXCEL.EXE en cours, à prendre juste avant <c>CreateInstance</c> pour
        /// retrouver ensuite le PID de notre instance (<see cref="ResoudrePidExcelCree"/>).
        /// Plus fiable que le HWND, absent tant que l'instance cachée n'a pas de fenêtre.
        /// </summary>
        private static HashSet<int> SnapshotPidsExcel()
        {
            var set = new HashSet<int>();
            try
            {
                foreach (var p in Process.GetProcessesByName("EXCEL"))
                {
                    try { set.Add(p.Id); } catch { /* process en train de mourir */ }
                }
            }
            catch { /* énumération impossible : tant pis, on renvoie un set vide */ }
            return set;
        }

        /// <summary>
        /// PID du EXCEL.EXE apparu depuis le snapshot <paramref name="avant"/> (notre instance COM),
        /// ou 0 si rien de neuf : l'appelant se rabat alors sur le HWND.
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

        /// <summary>
        /// Range les EXCEL.EXE externes en deux catégories :
        /// <c>visibles</c> (fenêtre principale : classeur de l'utilisateur, jamais fermé sans
        /// confirmation) et <c>fantomes</c> (sans fenêtre : reliquat de pilotage COM, fermable
        /// en silence pour libérer un verrou de fichier).
        /// </summary>
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
                catch { fantomes.Add(p); /* process en train de mourir : on le compte comme fantôme */ }
            }
            return (visibles, fantomes);
        }

        /// <summary>
        /// Vrai si un classeur de mesure est ouvert dans l'instance Excel.
        /// <see cref="MesureOrchestrator"/> en déduit son chemin : false (1ère mesure) → ClosedXML
        /// + OuvrirEtAfficher ; true → <see cref="AjouterFeuilleMesureAsync"/> en pur COM, sans
        /// fermer le classeur (évite la fenêtre grise SDI).
        /// </summary>
        public bool AClasseurActif
        {
            get { lock (_sync) { return _classeurActif != null; } }
        }

        /// <summary>Chemin du classeur actuellement ouvert dans Excel, ou vide s'il n'y en a aucun.</summary>
        public string CheminClasseurActif
        {
            get { lock (_sync) { return _cheminClasseurActif; } }
        }


        /// <summary>
        /// PID du EXCEL.EXE qu'on a démarré, pour le tuer en urgence si les appels COM se bloquent
        /// (Excel fermé à la main pendant une écriture : RPC suspendus ~60 s). 0 si non démarré.
        /// </summary>
        private int _excelPid;

        private readonly object _sync = new();
        private bool _disposed;

        // P/Invoke pour retrouver le PID du process Excel à partir de son handle de fenêtre.
        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // P/Invoke pour énumérer puis masquer les fenêtres top-level Excel parasites — la fameuse
        // fenêtre « shell » d'Excel SDI qui reste visible, toute grise, à côté du rapport.
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
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
        private const int SW_HIDE = 0;
        private const int SW_FORCEMINIMIZE = 11;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOACTIVATE = 0x0010;
        private static readonly IntPtr HWND_BOTTOM = new IntPtr(1);

        // Masque en continu les fenêtres shell (toutes les 500 ms tant qu'un classeur est ouvert),
        // même si la toute première tentative a raté à cause du timing COM.
        private System.Threading.Timer? _timerMasquageShell;
        private readonly object _timerLock = new();

        private ExcelInteropHost() { }

        /// <summary>Vrai si Excel est démarré et prêt à ouvrir un classeur.</summary>
        public bool EstDemarre => _excel != null;

        /// <summary>
        /// Vérifie au niveau OS si le process Excel (PID enregistré) est encore vivant.
        /// Ne prend PAS le lock <c>_sync</c> : appelable en boucle par un watchdog sans bloquer
        /// ni être bloqué par les appels COM en cours.
        /// </summary>
        public bool EstProcessExcelEnVie()
        {
            int pid = _excelPid;
            // PID inconnu (pas encore résolu, ou redémarrage) : on suppose vivant, sinon la
            // surveillance de l'accueil abandonnerait une mesure alors qu'Excel tourne.
            if (pid <= 0) return true;
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

                    // PID de NOTRE instance (diff des process, repli sur le HWND). Sert à s'exclure
                    // de la détection « Excel externe » et à TuerProcessExcelAsync() si le COM bloque.
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

                    // Pré-ouvre Metrologo.xla (comme le Delphi historique) : sans lui, les formules
                    // [1]!Cal_xxx des rapports affichent #NOM? tant que l'utilisateur n'a pas
                    // cliqué « Ouvrir le classeur » dans « Liaisons de classeur ».
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
        /// Ouvre le classeur sauvegardé par <c>ExcelService</c> dans l'instance cachée, active la
        /// feuille de mesure et affiche Excel. Un classeur déjà ouvert est fermé (une seule mesure
        /// live à la fois).
        /// </summary>
        public Task OuvrirEtAfficherAsync(string cheminFichier, string nomFeuilleMesure)
            => Task.Run(() =>
            {
                lock (_sync)
                {
                    // Excel a pu être fermé à la main ou planter : on relance une instance, sinon
                    // RPC déconnecté (0x800706BA) au prochain accès à _excel.
                    if (!EstInstanceVivante())
                    {
                        JournalLog.Warn(CategorieLog.Excel, "EXCEL_HOTE_REDEMARRE",
                            "Instance Excel hôte indisponible (fermée ou crashée) — redémarrage automatique.");
                        RedemarrerInstanceInterne();
                        if (_excel == null) return;   // redémarrage a échoué, on abandonne
                    }

                    // Fenêtre Excel fermée entre deux mesures : Excel ferme tous ses classeurs, .xla
                    // compris, mais le process survit. Sans ce contrôle : #NOM? sur [1]!Cal_xxx.
                    VerifierXlaChargeInterne();

                    // Ferme tout classeur précédent (une seule mesure live à la fois).
                    FermerClasseurActifInterne();

                    // UpdateLinks=0 : le .xla est déjà pré-ouvert (PreChargerXlaInterne), donc Excel
                    // résout [1]!Cal_xxx sans re-scanner les liens (UpdateLinks=3 = recherche réseau lente).
                    _classeurActif = _excel.Workbooks.Open(cheminFichier, 0);
                    _cheminClasseurActif = cheminFichier;
                    try { _fullNameActifExcel = (string)_classeurActif.FullName; }
                    catch { _fullNameActifExcel = string.Empty; }

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

                    // Recalcul forcé : sans lui le graphe Stab (axe Y log) reste vide, car ClosedXML ne
                    // peuple pas le numCache des charts et UpdateLinks=0 ne déclenche pas de recalcul.
                    try { _excel.CalculateFullRebuild(); }
                    catch (Exception ex)
                    {
                        JournalLog.Warn(CategorieLog.Excel, "EXCEL_RECALC_KO",
                            $"CalculateFullRebuild échoué : {ex.Message}");
                    }

                    // Ferme les classeurs résiduels (Book1 vide, fichiers ouverts par double-clic) qui
                    // laisseraient des fenêtres Excel grises à côté du classeur actif.
                    FermerClasseursParasitesInterne();

                    _excel.ScreenUpdating = true;
                    _excel.Visible = true;

                    // Élargit la zone des onglets (défaut ~60 %) à 85 % de la barre, pour voir d'un
                    // coup les onglets 1..13 d'une stab.
                    try
                    {
                        dynamic w = _classeurActif.Windows[1];
                        try { w.TabRatio = 0.85; }
                        finally { try { Marshal.ReleaseComObject(w); } catch { } }
                    }
                    catch { /* best-effort, non bloquant */ }

                    // Masque la shell SDI vide qui apparaît avec Visible=true (léger clignotement possible).
                    MasquerFenetresShellExcelInterne();
                    MasquerFenetresNonActivesInterne();

                    // Rapport au premier plan (annule une éventuelle shell TopMost créée par Workbooks.Open).
                    ForcerRapportAuPremierPlanInterne();

                    // Timer de masquage continu : re-masque une shell qui réapparaîtrait (Save, Open...).
                    DemarrerTimerMasquageShellInterne();

                    JournalLog.Info(CategorieLog.Excel, "EXCEL_CLASSEUR_OUVERT",
                        $"Classeur ouvert dans Excel : {Path.GetFileName(cheminFichier)} → feuille {nomFeuilleMesure}.");
                }
            });

        /// <summary>
        /// Écrit la i-ème mesure (heure + valeur brute, colonnes A:B). Les formules des colonnes C
        /// (<c>Fréq. Réelle</c>) et D (<c>F(i)-F(i+1)</c>) sont déjà posées par l'init ClosedXML.
        /// </summary>
        /// <param name="indexMesure">0-based : la 1ère mesure va en <paramref name="ligneDebut"/>.</param>
        /// <param name="ligneDebut">Première ligne de la zone de mesures (9 dans le template).</param>
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
                        // Écriture en bloc A:B (un seul Range.Value2) plutôt que 2 Cells[r,c] : une seule
                        // réactivation du rectangle de sélection (moins de clignotement), ~30-40 % plus rapide.
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
        /// Écrit toutes les mesures d'une gate en un seul appel COM, bien plus rapide que N appels
        /// <see cref="EcrireValeurLiveAsync"/> quand le suivi live n'a pas d'intérêt (balayage de
        /// stabilité : 30 mesures à ~10 ms par gate).
        /// </summary>
        /// <param name="ligneDebut">Première ligne de la zone (9 dans le template).</param>
        /// <param name="mesures">Liste ordonnée (timestamp, valeur) ; l'index 0 va en <paramref name="ligneDebut"/>.</param>
        public Task EcrireValeursEnBlocAsync(int ligneDebut, IList<(DateTime ts, double valeur)> mesures)
            => Task.Run(() =>
            {
                lock (_sync)
                {
                    var feuille = _feuilleMesure;
                    if (feuille == null || mesures.Count == 0) return;
                    try
                    {
                        // Matrice N lignes × 2 colonnes (heure, mesure), 0-based pour Range.Value2.
                        object[,] matrix = new object[mesures.Count, 2];
                        for (int i = 0; i < mesures.Count; i++)
                        {
                            matrix[i, 0] = mesures[i].ts.ToString("HH:mm:ss");
                            matrix[i, 1] = mesures[i].valeur;
                        }

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

        // ---------- Format des rapports : .xls en sortie de mesure, .xlsx pour travailler ----------

        private const int FormatExcelXls = 56;    // xlExcel8 (Excel 97-2003)
        private const int FormatExcelXlsx = 51;   // xlOpenXMLWorkbook

        /// <summary>
        /// Sortie de mesure en mode .xls : enregistre le rapport <paramref name="cheminXlsx"/> en .xls
        /// au même endroit (même nom), puis supprime le .xlsx — il ne reste qu'un fichier Excel
        /// dans le dossier de la FI. Si c'est le classeur ouvert, il DEVIENT le .xls : les saisies
        /// post-mesure et les retouches manuelles y sont enregistrées directement. Renvoie false si
        /// la conversion a échoué (ex. .xls ouvert ailleurs) : le .xlsx est alors conservé.
        /// </summary>
        public Task<bool> ConvertirRapportEnXlsAsync(string cheminXlsx) => Task.Run(() =>
        {
            lock (_sync)
            {
                return ConvertirRapportInterne(cheminXlsx, Path.ChangeExtension(cheminXlsx, ".xls"), FormatExcelXls);
            }
        });

        /// <summary>
        /// Avant une mesure : si le rapport de la FI n'existe qu'en .xls (sortie de la mesure
        /// précédente), le repasse en .xlsx (<paramref name="cheminXlsx"/>) et supprime le .xls —
        /// le moteur des rapports (ClosedXML) ne travaille qu'en .xlsx. Sans .xls : ne fait rien.
        /// </summary>
        public Task<bool> ReconvertirRapportEnXlsxAsync(string cheminXlsx) => Task.Run(() =>
        {
            lock (_sync)
            {
                string cheminXls = Path.ChangeExtension(cheminXlsx, ".xls");
                if (!File.Exists(cheminXls)) return true;
                if (File.Exists(cheminXlsx))
                {
                    // Les deux coexistent (conversion interrompue) : on garde le plus récent.
                    if (File.GetLastWriteTime(cheminXlsx) >= File.GetLastWriteTime(cheminXls))
                    {
                        if (!MemeCheminFichier(_cheminClasseurActif, cheminXls)) SupprimerFichierAvecReessais(cheminXls);
                        return true;
                    }
                }
                return ConvertirRapportInterne(cheminXls, cheminXlsx, FormatExcelXlsx);
            }
        });

        /// <summary>Convertit <paramref name="source"/> en <paramref name="cible"/> (format Excel
        /// <paramref name="format"/>) puis supprime la source. Classeur ouvert : SaveAs dessus (il
        /// devient la cible). Sinon : instance Excel invisible et éphémère, pour ne rien changer à
        /// ce que l'utilisateur voit (cas typique : Excel tué par le bouton STOP). Sous lock.</summary>
        private bool ConvertirRapportInterne(string source, string cible, int format)
        {
            if (!File.Exists(source)) return false;
            try
            {
                if (_classeurActif != null && MemeCheminFichier(_cheminClasseurActif, source))
                {
                    try { _excel.DisplayAlerts = false; } catch { }
                    try { _classeurActif.CheckCompatibility = false; } catch { }
                    _classeurActif.SaveAs(cible, format);
                    _cheminClasseurActif = cible;
                    try { _fullNameActifExcel = (string)_classeurActif.FullName; } catch { }
                }
                else
                {
                    ConvertirDansInstanceSepareeInterne(source, cible, format);
                }

                SupprimerFichierAvecReessais(source);
                JournalLog.Info(CategorieLog.Excel, "RAPPORT_CONVERTI",
                    $"Rapport converti : {Path.GetFileName(source)} → {Path.GetFileName(cible)} ({Path.GetDirectoryName(cible)}).");
                return true;
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "RAPPORT_CONVERSION_KO",
                    $"Conversion {Path.GetFileName(source)} → {Path.GetFileName(cible)} échouée : {ex.Message}");
                return false;
            }
        }

        private static void ConvertirDansInstanceSepareeInterne(string source, string cible, int format)
        {
            dynamic? app = null;
            dynamic? classeurs = null;
            dynamic? wb = null;
            try
            {
                app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")
                    ?? throw new InvalidOperationException("Excel n'est pas installé."))!;
                try { app.Visible = false; } catch { }
                try { app.DisplayAlerts = false; } catch { }
                try { app.AskToUpdateLinks = false; } catch { }
                try { app.EnableEvents = false; } catch { }
                classeurs = app.Workbooks;
                wb = classeurs.Open(source, 0);            // UpdateLinks = 0 : pas de question sur le .xla
                try { wb.CheckCompatibility = false; } catch { }
                wb.SaveAs(cible, format);
                wb.Close(false);
            }
            finally
            {
                if (wb != null) { try { Marshal.ReleaseComObject(wb); } catch { } }
                if (classeurs != null) { try { Marshal.ReleaseComObject(classeurs); } catch { } }
                if (app != null)
                {
                    try { app.Quit(); } catch { }
                    try { Marshal.ReleaseComObject(app); } catch { }
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static void SupprimerFichierAvecReessais(string chemin)
        {
            for (int i = 0; i < 10 && File.Exists(chemin); i++)
            {
                try { File.Delete(chemin); }
                catch { System.Threading.Thread.Sleep(200); }   // Excel relâche son handle un peu après
            }
            if (File.Exists(chemin))
                JournalLog.Warn(CategorieLog.Excel, "RAPPORT_ANCIEN_FORMAT_KO",
                    $"Ancien fichier non supprimé (verrouillé ?) : {chemin}");
        }

        private static bool MemeCheminFichier(string? a, string? b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return false;
            try { return string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase); }
            catch { return false; }
        }

        /// <summary>
        /// Écrit une valeur dans une zone nommée (portée feuille) du classeur actif, pour les
        /// saisies post-mesure (fréquence lue, incertitudes).
        /// </summary>
        public Task EcrireZoneNommeeAsync(string zoneNom, object valeur) => Task.Run(() =>
        {
            lock (_sync)
            {
                if (_feuilleMesure == null) return;

                // Les feuilles sont protégées en fin de mesure (ProtegerToutesFeuillesActifInterne) :
                // on déprotège avant d'écrire, sinon Value2 lève et la valeur est perdue.
                bool etaitProtegee = false;
                try { etaitProtegee = (bool)_feuilleMesure.ProtectContents; } catch { }
                if (etaitProtegee)
                {
                    try { _feuilleMesure.Unprotect(MotDePasseProtectionFeuille); }
                    catch { try { _feuilleMesure.Unprotect("metrol"); } catch { try { _feuilleMesure.Unprotect(); } catch { } } }
                }

                try
                {
                    dynamic nom = _feuilleMesure.Names.Item(zoneNom);
                    try { nom.RefersToRange.Value2 = valeur; }
                    finally { try { Marshal.ReleaseComObject(nom); } catch { } }
                }
                catch (Exception ex)
                {
                    // Zone absente ou erreur COM : on loggue en Warn, pas d'interruption.
                    JournalLog.Warn(CategorieLog.Excel, "EXCEL_ZONE_ECRITURE_ERREUR",
                        $"Impossible d'écrire la zone nommée « {zoneNom} » : {ex.Message}");
                }
                finally
                {
                    if (etaitProtegee)
                    {
                        try { _feuilleMesure.Protect(MotDePasseProtectionFeuille); } catch { }
                    }
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
        /// Urgence : tue le process EXCEL.EXE par son PID, sans <c>_sync</c> ni appel COM, quand un
        /// appel COM est bloqué (Excel fermé pendant une écriture : RPC en attente ~60 s). Les appels
        /// bloqués lèvent aussitôt <c>COMException 0x800706BA</c> et libèrent <c>_sync</c> ; l'instance
        /// devra être recréée (<see cref="EstInstanceVivante"/> + <see cref="RedemarrerInstanceInterne"/>).
        /// </summary>
        public Task TuerProcessExcelAsync() => Task.Run(() =>
        {
            // Volontairement SANS lock(_sync) : on veut débloquer les threads qui le tiennent.
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
                // Déjà mort : c'est le résultat voulu.
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
                // Marque l'instance comme morte.
                _excelPid = 0;
                _excel = null;
                _classeurActif = null;
                _feuilleMesure = null;
                _cheminClasseurActif = string.Empty;
                _fullNameActifExcel = string.Empty;
            }
        });

        /// <summary>
        /// Vérifie que l'instance répond encore au COM, avant chaque ouverture de classeur
        /// (évite le RPC déconnecté 0x800706BA si Excel a été fermé ou a crashé).
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
        /// Recrée une instance Excel après perte de la précédente (fermée, crashée).
        /// À appeler sous lock <see cref="_sync"/>.
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
            _fullNameActifExcel = string.Empty;
            _cheminXlaPreCharge = string.Empty;   // le .xla appartenait à l'ancienne instance
            // On oublie l'ancien PID tout de suite : conservé pendant le redémarrage (plusieurs
            // secondes), il ferait conclure à la surveillance de l'accueil qu'Excel est mort
            // et abandonner la mesure.
            _excelPid = 0;

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

                // PID du NOUVEAU process AVANT le pré-chargement du .xla : sinon TuerProcessExcelAsync
                // viserait un PID périmé et le nettoyage des Excel « fantômes » pourrait fermer
                // notre propre instance.
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

                PreChargerXlaInterne();
                JournalLog.Info(CategorieLog.Excel, "EXCEL_HOTE_REDEMARRE",
                    $"Instance Excel recréée (PID={_excelPid}).");
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
        /// Rouvre Metrologo.xla s'il n'est plus ouvert dans l'instance (testé par nom : un
        /// complément n'apparaît pas dans l'énumération de Workbooks mais reste accessible par
        /// <c>Workbooks.Item(nom)</c>). À appeler sous lock <see cref="_sync"/>.
        /// </summary>
        private void VerifierXlaChargeInterne()
        {
            if (_excel == null) return;
            string nomXla = Path.GetFileName(Preferences.CheminMacroXLA ?? string.Empty);
            if (string.IsNullOrEmpty(nomXla)) return;
            try
            {
                dynamic wb = _excel.Workbooks.Item(nomXla);
                try { Marshal.ReleaseComObject(wb); } catch { }
                return;   // toujours là
            }
            catch { /* absent : on le recharge */ }

            JournalLog.Warn(CategorieLog.Excel, "XLA_RECHARGE",
                "Metrologo.xla n'était plus ouvert dans l'instance Excel (fenêtre Excel fermée entre "
                + "deux mesures ?) — rechargement avant d'ouvrir le rapport.");
            _cheminXlaPreCharge = string.Empty;
            PreChargerXlaInterne();
        }

        /// <summary>
        /// Ouvre Metrologo.xla comme classeur pour que les formules <c>[1]!Cal_xxx(...)</c> des
        /// rapports se résolvent sans le panneau « Liaisons de classeur » (comme le Delphi, cf.
        /// F_Main.pas : XLApp.Workbooks.Open(S_XLA, ...)). Un add-in chargé via HKCU\Excel\Options\OPENn
        /// ne suffit pas : il fournit <c>=Cal_xxx(...)</c> mais pas <c>=[1]!Cal_xxx(...)</c>.
        /// À appeler sous lock <see cref="_sync"/>, <see cref="_excel"/> non null.
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
                // UpdateLinks=0 (pas de liens externes), ReadOnly=true (on n'altère pas le .xla).
                dynamic wb = _excel.Workbooks.Open(xlaPath, 0, true);
                _cheminXlaPreCharge = (string)wb.FullName;
                // IsAddin=true masque sa fenêtre : sinon Excel 365 le traite en classeur normal et
                // affiche une fenêtre grise vide à côté du rapport (Visible=true). Il reste résolvable par les liens.
                try { wb.IsAddin = true; } catch { /* souvent déjà true */ }
                // Excel garde le classeur via Workbooks : on libère juste le proxy COM.
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
        /// Masque toutes les fenêtres Excel SAUF celle du classeur actif : Excel réaffiche celles du
        /// .xla et des classeurs résiduels à chaque toggle <c>Visible=false→true</c>, d'où une 2e
        /// fenêtre grise dès la 2e mesure. À appeler sous lock <see cref="_sync"/>, après <c>Visible=true</c>.
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
                            bool estClasseurActif = EstClasseurActif(fullName);
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
        /// Ferme les classeurs de mesure résiduels de l'instance hôte (Book1, anciens rapports) SAUF
        /// <see cref="_classeurActif"/> et le .xla (<see cref="_cheminXlaPreCharge"/>), pour éviter les
        /// fenêtres grises parasites. À appeler sous lock <see cref="_sync"/>.
        /// </summary>
        private void FermerClasseursParasitesInterne()
        {
            if (_excel == null) return;
            try
            {
                // Index décroissant : la collection change à chaque Close (un foreach lèverait).
                int nb = _excel.Workbooks.Count;
                var diag = new System.Text.StringBuilder($"actif=[{_cheminClasseurActif}] xla=[{_cheminXlaPreCharge}] nb={nb} ");
                for (int i = nb; i >= 1; i--)
                {
                    dynamic wb = _excel.Workbooks[i];
                    try
                    {
                        string fullName = (string)wb.FullName;
                        bool estClasseurActif = EstClasseurActif(fullName);
                        bool estXlaPreCharge = !string.IsNullOrEmpty(_cheminXlaPreCharge)
                            && ChemPathEgal(fullName, _cheminXlaPreCharge);
                        // Seuls les classeurs de mesure sont fermés : un fichier personnel de
                        // l'utilisateur ne doit jamais l'être (Close(false) = sans sauvegarde).
                        bool estClasseurMesure = EstClasseurDeMesure(fullName);
                        string action = "skip";
                        if (!estClasseurActif && !estXlaPreCharge && estClasseurMesure)
                        {
                            try { wb.Close(false); action = "FERME"; }
                            catch { action = "ECHEC_CLOSE"; }
                        }
                        else if (!estClasseurActif && !estXlaPreCharge && !estClasseurMesure)
                        {
                            action = "skip(perso)";
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

        /// <summary>
        /// Vrai si le classeur est un rapport de mesure Metrologo : nom <c>freq</c> / <c>stab</c>
        /// (suffixe numérique de fallback possible : freq2, stab3…) ET dossier de mesures
        /// (<c>…\Metrologo\&lt;FI&gt;\</c> ou <see cref="CheminsMetrologo.MesuresLocal"/>, que
        /// Excel peut aussi écrire en <c>\\serveur\partage\…</c>). La double condition évite de
        /// refermer un fichier personnel au nom proche.
        /// </summary>
        private static bool EstClasseurDeMesure(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName)) return false;
            try
            {
                string nom = Path.GetFileNameWithoutExtension(fullName);
                bool nomMesure = System.Text.RegularExpressions.Regex.IsMatch(
                    nom, @"^(freq|stab)\d*$",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                bool dansMetrologo =
                    fullName.IndexOf(@"\Metrologo\", StringComparison.OrdinalIgnoreCase) >= 0
                    || EstSousDossier(fullName, CheminsMetrologo.MesuresLocal);
                return nomMesure && dansMetrologo;
            }
            catch { return false; }
        }

        // ---------- Finalisation de mesure 100 % en COM ----------
        // On écrit directement dans le classeur ouvert, sans cycle ferme/écrit-ClosedXML/rouvre
        // (cause d'une fenêtre grise parasite en fin de mesure). Couvre : coefficients
        // d'incertitude (ZNCoeffA/B/C/D), ligne Récap (Freq), Save. Stab et Tachy (fréquence
        // auxiliaire) passent par le même chemin ; le graphe Stab reste créé via ClosedXML.

        /// <summary>
        /// Finalise la mesure via Excel COM : coefficients d'incertitude du module sélectionné,
        /// ligne Récap du type de mesure, puis Save du classeur actif (qui reste visible).
        /// </summary>
        public Task FinaliserMesureViaComAsync(Mesure mesure, List<double> resultats,
            string nomFeuille, double tempsGateSecondes, bool isDerniereGate = false,
            int nbGatesStabBalayees = 0) => Task.Run(() =>
        {
            lock (_sync)
            {
                if (_excel == null || _classeurActif == null) return;

                // Rendu Excel gelé pendant la finalisation : sinon les ~80 opérations COM sur la
                // feuille Récap provoquent chacune un repaint (effet de « balayage » visible).
                bool screenUpdatingOriginal = true;
                int calculationOriginal = -4105;   // xlCalculationAutomatic
                try { screenUpdatingOriginal = (bool)_excel.ScreenUpdating; } catch { }
                try { calculationOriginal = (int)_excel.Calculation; } catch { }
                try { _excel.ScreenUpdating = false; } catch { }
                try { _excel.Calculation = -4135; /* xlCalculationManual */ } catch { }

                try
                {
                    // La formule de variance (ZNVariance, `[1]!Cal_variance(SUMSQ(G10:G38),...)`) est
                    // posée par ClosedXML (PreparerLignesMesureAsync) : ne pas l'écraser par une valeur,
                    // elle permet le recalcul si l'utilisateur modifie les données.

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

                        // Dernière gate Stab : plages des séries du graphe via COM (équivalent de
                        // RendreGrapheStabDynamique de ExcelService). Bornes de l'axe Y : après
                        // le recalcul (étape 4), car elles se calent sur les valeurs de la Récap.
                        if (isDerniereGate)
                        {
                            int nbGates = nbGatesStabBalayees > 0
                                ? nbGatesStabBalayees
                                : _sigmasRelatifsParGate.Count;
                            AjusterPlagesGrapheStabViaComInterne(nbGates);
                        }
                    }

                    // 3. Redirige le lien externe Metrologo.xla vers Preferences.CheminMacroXLA (ChangeLink
                    //    COM), pour qu'Excel ne demande pas de repointer le .xla si le fichier
                    //    contient des chemins obsolètes.
                    RedirigerLinksXlaViaComInterne();

                    // 4. Recalcul complet AVANT le Save. Les formules ont été posées en Calculation=Manual :
                    //    sans recalcul, le fichier n'a pas de valeurs en cache et, à la réouverture,
                    //    les cellules liées à Metrologo.xla (=[1]!Cal_variance...) s'affichent
                    //    « périmées » (barrées), le lien externe étant désactivé par défaut.
                    try { _excel.Calculation = -4105; /* xlCalculationAutomatic */ } catch { }
                    try { _classeurActif.Calculate(); } catch { }

                    if (mesure.TypeMesure == TypeMesure.Stabilite && isDerniereGate)
                    {
                        int nbGatesAxe = nbGatesStabBalayees > 0
                            ? nbGatesStabBalayees
                            : _sigmasRelatifsParGate.Count;
                        CalibrerAxeYGrapheStabViaComInterne(nbGatesAxe, _sigmasRelatifsParGate.Values);
                    }

                    // Pas d'actualisation des liens externes (Metrologo.xla) à la prochaine ouverture :
                    // sinon Excel marque le classeur « modifié » et propose d'enregistrer à la fermeture.
                    // Les valeurs en cache posées par le Calculate ci-dessus sont utilisées telles quelles.
                    try { _classeurActif.UpdateLinks = 2; /* xlUpdateLinksNever */ } catch { }

                    // Verrouillage post-mesure : à la dernière gate, on re-protège toutes les feuilles
                    // AVANT le Save (pour persister la protection). L'app les déprotège seule à la relance.
                    if (isDerniereGate)
                        ProtegerToutesFeuillesActifInterne();

                    // 5. Save
                    try { _classeurActif.Save(); }
                    catch (Exception ex)
                    {
                        JournalLog.Warn(CategorieLog.Excel, "EXCEL_SAVE_COM_KO",
                            $"Save COM final échoué : {ex.Message}");
                        // Rapport sur le réseau et réseau perdu en cours de série : on ne perd pas
                        // la mesure, on l'enregistre en local (rapatriée dès le retour du réseau).
                        EnregistrerEnLocalApresEchecInterne(mesure);
                    }

                    // Classeur « non modifié » après le Save : évite le prompt « Voulez-vous enregistrer
                    // les modifications ? » en cas de recalcul interne d'Excel.
                    try { _classeurActif.Saved = true; } catch { }

                    // 6. Active la feuille Récap. en fin de mesure ; pour Stab, seulement à la dernière
                    //    gate (sinon on ne verrait pas la feuille en cours entre deux gates).
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
                    // Restauration des flags Excel, dans cet ordre :
                    //  1. Calculation d'origine ;
                    //  2. Calculate() forcé : Excel ne recalcule pas seul les formules posées en mode
                    //     Manual, et les cellules figées (#VALEUR!, #DIV/0!, 0) garderaient le style
                    //     « erreur » du template (gris barré) ;
                    //  3. ScreenUpdating d'origine : Excel dessine l'état final d'un coup ;
                    //  4. Saved = true en dernier : le Calculate précédent a re-marqué le classeur
                    //     modifié, ce qui ferait proposer l'enregistrement à la fermeture.
                    try { _excel.Calculation = calculationOriginal; } catch { }
                    try { _classeurActif.Calculate(); } catch { }
                    try { _excel.ScreenUpdating = screenUpdatingOriginal; } catch { }
                    try { _classeurActif.Saved = true; } catch { }
                }
            }
        });

        /// <summary>
        /// Ajoute une feuille de mesure (copie de ModFeuille) au classeur ouvert, en pur COM et sans
        /// le fermer : l'ancienne séquence fermer/ClosedXML/rouvrir laissait Excel visible sans
        /// classeur pendant 1-3 s (shell SDI grise). Reproduit le travail de
        /// <c>ExcelService.InitialiserRapportAsync</c> + <c>PreparerLignesMesureAsync</c> :
        /// copie de la feuille, renommage/déprotection, <b>clonage des Names portée feuille</b>,
        /// en-têtes, métadonnées en zones nommées, formules, Activate + Save.
        /// </summary>
        /// <returns>Nom de la nouvelle feuille (ex : <c>"freq2"</c>, <c>"stab3"</c>).</returns>
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

                    // Rendu gelé pendant la séquence COM : ~6 s sans (33 Names.Add + 60 écritures de
                    // formules avec redraw), ~1-2 s avec. ForcerRapportAuPremierPlanInterne rafraîchit à la fin.
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

                        // Feuilles à nom fixe (avinter / fqfinale) : à la relance sur la même FI le nom
                        // existe déjà, on écrase l'ancienne feuille ET sa ligne Récap avant la copie
                        // (sinon le rename lèverait et la Récap aurait deux lignes). No-op pour les
                        // feuilles numérotées.
                        SupprimerFeuilleEtRecapInterne(nomNouvelleFeuille);

                        // --- 3. Copie ModFeuille à la fin du classeur ---
                        // Piège Excel COM : sans Before ni After, Copy crée un NOUVEAU classeur. En
                        // late-binding, Type.Missing pour Before était pris pour « non spécifié » (classeur
                        // parasite à chaque appel) : on passe donc le paramètre nommé After:.
                        dynamic derniereFeuille = _classeurActif.Worksheets
                            [(int)_classeurActif.Worksheets.Count];
                        try
                        {
                            modFeuille.Copy(After: derniereFeuille);
                        }
                        finally { try { Marshal.ReleaseComObject(derniereFeuille); } catch { } }

                        // La copie devient ActiveSheet du classeur cible (on lit celui du classeur, pas
                        // Application.ActiveSheet, qui peut pointer ailleurs).
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

                        // --- 5. Clonage des Names portée feuille (critique) ---
                        // Les Names des templates sont portée classeur et pointent vers ModFeuille!
                        // (33 sur METROLOGO.xlsx, idem Stab/Tachy) ; Worksheet.Copy ne les clone pas.
                        // On les recrée portée feuille, sinon ZNFreqMoyReel, ZNVariance, ZNEcartType...
                        // résoudraient vers ModFeuille (vide) au lieu de la feuille de mesure.
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
                                // Worksheet.Names.Add (portée feuille), pas Workbook.Names.Add (conflit avec l'existant).
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

                        // --- 7. Module d'incertitude sélectionné (config.NumModuleIncertitude), écrit une
                        // seule fois à côté du label « Module = » : F10 (standard) ou G10 (tachy, 1 colonne de plus).
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
                        // Résolution et incertitude relative à 0 : seule la saisie post-mesure les renseigne.
                        // Ne PAS seeder ZNIncertSup depuis config.IncertSupp (config est réutilisé entre
                        // relances sur une même FI : la valeur saisie au fréquencemètre réapparaîtrait sans
                        // popup sur la mesure suivante). Cf. ExcelService.InitialiserRapportAsync.
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNIncertResol", 0.0);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNIncertSup", 0.0);
                        // Fréquence de référence (rubidium actif) écrite DIRECTEMENT en C21 (ZNFreqRef pointe
                        // dessus) : la résolution par zone nommée pouvait retomber sur la cellule de
                        // ModFeuille (cachée) et laisser C21 vide sur la feuille visible.
                        EcrireValeurCelluleInterne(nouvelleFeuille, "C21", rubidium.FrequenceMoyenne);
                        // 7 décimales : maximum fidèle à ~10 MHz (Excel garde 15 chiffres significatifs =
                        // 8 entiers + 7 décimales) ; au-delà, des « 00 » parasites apparaissent.
                        try
                        {
                            dynamic cFmt = nouvelleFeuille.Range["C21"];
                            try { cFmt.NumberFormat = "#,##0.0000000"; }
                            finally { try { Marshal.ReleaseComObject(cFmt); } catch { } }
                        }
                        catch { /* best-effort */ }
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNFreqRef",
                            rubidium.FrequenceMoyenne);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNCoeffA", 1e-10);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNCoeffB", 5e-13);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNNbMesAccredite", 30.0);
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNTempsMesureAccredite", 10.0);

                        // --- 10. Lignes additionnelles (PreparerLignesMesure) ---
                        // Le template a 2 lignes (9 et 10) : pour nbMesures > 2, on insère
                        // (nbMesures - 2) rangées avant ZNPointInsertion (défaut : ligne 11).
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

                        // --- 11. Formules C (conversion), D (delta) et K (tr/min, tachy) des lignes 11..N ---
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

                        // --- 12. Formules métier (ZNFreqMoyReel, ZNVariance, ZNEcartType, …) ---
                        // Posées via Names → RefersToRange.Formula (Names portée feuille de l'étape 5 ;
                        // F13 = moyenne, F18 = variance, F19 = écart-type... selon le layout ModFeuille).
                        int ligneDeb = 9;
                        int ligneFin = 9 + nbMesures - 1;
                        EcrireFormuleZoneNommeeInterne(nomNouvelleFeuille, "ZNFreqMoyReel",
                            $"=IF(ISBLANK(ZNNbMesures),,AVERAGE(C{ligneDeb}:C{ligneFin}))");

                        // Piège Excel COM : `=[1]!Cal_variance(SUMSQ(...),...)` est rejetée (0x800A03EC)
                        // par .Formula/.Formula2/.FormulaLocal en COM dynamic (elle passe via ClosedXML),
                        // et les formules `=IF(ISBLANK(ZNVariance),...)` sont rejetées si ZNVariance
                        // pointe sur une cellule totalement vide. D'où un placeholder 0 AVANT de poser
                        // les formules qui en dépendent ; la variance réelle le remplace ensuite.
                        EcrireValeurZoneNommeeInterne(nomNouvelleFeuille, "ZNVariance", 0.0);

                        EcrireFormuleZoneNommeeInterne(nomNouvelleFeuille, "ZNEcartType",
                            "=IF(ISBLANK(ZNVariance),,[1]!Cal_ecart_type(ZNVariance))");
                        EcrireFormuleZoneNommeeInterne(nomNouvelleFeuille, "ZNIncertEcartType",
                            "=IF(ISBLANK(ZNNbMesures),,[1]!Cal_incert_ecart_type(ZNVariance,ZNNbMesures))");
                        EcrireFormuleZoneNommeeInterne(nomNouvelleFeuille, "ZNFreqCorr",
                            "=IF(ISBLANK(ZNFreqMoyReel),,[1]!Cal_freq_corrigee(ZNFreqMoyReel,ZNFreqRef))");

                        // --- 13. Format General sur les cellules de formules : en format Texte hérité,
                        // Excel garderait la formule littérale (cf. EcrireFormuleRecapInterne).
                        try
                        {
                            dynamic rangeFormules = nouvelleFeuille.Range[$"C{ligneDeb}:D{ligneFin}"];
                            try { rangeFormules.NumberFormat = "General"; }
                            finally { try { Marshal.ReleaseComObject(rangeFormules); } catch { } }
                        }
                        catch { }

                        // --- 14. Active la feuille ---
                        nouvelleFeuille.Activate();

                        // Activate peut réveiller le window manager Excel : on re-masque les shells.
                        MasquerFenetresShellExcelInterne();
                        DemarrerTimerMasquageShellInterne();

                        // Feuille de mesure courante
                        if (_feuilleMesure != null)
                        {
                            try { Marshal.ReleaseComObject(_feuilleMesure); } catch { }
                        }
                        _feuilleMesure = nouvelleFeuille;

                        // --- 15. Save ---
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

                        // Ceinture et bretelles : ferme les classeurs parasites qu'un Copy aurait créés
                        // malgré After: (no-op en pratique).
                        try { FermerClasseursParasitesInterne(); } catch { }

                        // Restauration des flags : Calculation d'abord (recalcule les formules posées),
                        // ScreenUpdating en dernier (rendu final d'un coup, sans flash).
                        try { _excel.Calculation = calculationOriginal; } catch { }
                        try { _excel.EnableEvents = enableEventsOriginal; } catch { }
                        try { _excel.ScreenUpdating = screenUpdatingOriginal; } catch { }

                        try { MasquerFenetresShellExcelInterne(); } catch { }

                        // Indispensable : sans cela, Excel SDI peut laisser une fenêtre grise TopMost
                        // par-dessus le rapport (elle disparaît au clic sur l'autre Excel).
                        try { ForcerRapportAuPremierPlanInterne(); } catch { }
                    }
                }
            });

        /// <summary>
        /// Réinitialise le classeur ouvert (« fichier vierge ») en supprimant toutes les feuilles de
        /// mesure (freq*, stab*, inter*, topti*, tcont*, strob*, avinter, fqfinale) sauf ModFeuille et
        /// Récap., SANS fermer le classeur : évite la shell SDI de « Relancer » → « Oui = écraser ».
        /// La Récap. est aussi nettoyée, ses formules vers les feuilles supprimées devenant #REF!.
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

                    // Freeze léger (cf. AjouterFeuilleMesureAsync) : pas de ScreenUpdating=false,
                    // qui donnerait un « écran gris » pendant la suppression des feuilles.
                    bool enableEventsOriginal = true;
                    int calculationOriginal = -4105;
                    try { enableEventsOriginal = (bool)_excel.EnableEvents; } catch { }
                    try { calculationOriginal = (int)_excel.Calculation; } catch { }
                    try { _excel.EnableEvents = false; } catch { }
                    try { _excel.DisplayAlerts = false; } catch { }
                    try { _excel.Calculation = -4135; } catch { }

                    try
                    {
                        // 1. Feuilles à supprimer (toutes sauf ModFeuille et Récap.) : noms collectés
                        //    d'abord, pour ne pas muter la collection pendant l'itération.
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
                                        // Conservées : ModFeuille (modèle de copie) et Récap. ; les autres
                                        // (freq1, stab1, ... et anciennes "1".."10" résiduelles) sautent.
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

                        // 3. Nettoyage Récap. : lignes de données à partir de la ligne 6 (entête en 5 pour
                        //    la Récap. Fréquence ; large pour couvrir Stab), sinon les formules
                        //    `='freq1'!ZN*` deviennent #REF!.
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

                                // Plages nommées du récap ramenées à une ligne (récap vide). Pas
                                // pour la stabilité, dont le récap est encore exploité à la main.
                                string nomClasseur = "";
                                try { nomClasseur = (string)_classeurActif.Name; } catch { }
                                if (!nomClasseur.StartsWith("stab", StringComparison.OrdinalIgnoreCase))
                                    MettreAJourPlagesNommeesRecapInterne(recap);
                            }
                            finally { try { Marshal.ReleaseComObject(recap); } catch { } }
                        }
                        catch { /* Récap. absent : pas de souci */ }

                        // 4. Save : persiste l'état vierge (ModFeuille + Récap. propre) sous le même chemin.
                        try { _classeurActif.Save(); }
                        catch (Exception ex)
                        {
                            JournalLog.Warn(CategorieLog.Excel, "REINIT_CLASSEUR_SAVE_KO",
                                $"Save après nettoyage échoué : {ex.Message}");
                        }

                        // _feuilleMesure pointait sur une feuille supprimée : la prochaine
                        // AjouterFeuilleMesureAsync recréera freq1.
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
        /// Supprime la feuille <paramref name="nomFeuille"/> (et sa ligne Récap.) puis sauvegarde.
        /// Appelé à l'arrêt d'une mesure : la feuille créée au démarrage n'est pas conservée.
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

                    // La feuille courante vient d'être supprimée : on lâche la référence COM.
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
        /// Supprime la feuille <paramref name="nomFeuille"/> du classeur actif si elle existe, ainsi
        /// que les lignes de la Récap. qui la référencent (DisplayAlerts coupé pour éviter la
        /// confirmation Excel). No-op si elle n'existe pas. À appeler sous lock <c>_sync</c>.
        /// </summary>
        private void SupprimerFeuilleEtRecapInterne(string nomFeuille)
        {
            if (_classeurActif == null || string.IsNullOrWhiteSpace(nomFeuille)) return;

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

        /// <summary>
        /// Nom de feuille unique (freq2, stab3, inter1, …) d'après les feuilles du classeur actif.
        /// Reproduit <c>ExcelService.TrouverNomFeuilleUnique</c>.
        /// </summary>
        private string TrouverNomFeuilleUniqueInterne(TypeMesure type)
        {
            if (_classeurActif == null) return "1";
            if (type == TypeMesure.FreqAvantInterv) return "F_avant_interv";
            if (type == TypeMesure.FreqFinale) return "F_finale";

            string prefixe = type switch
            {
                TypeMesure.Frequence    => "",   // numéros seuls (1, 2, 3, …)
                TypeMesure.Stabilite    => "",   // 1 feuille par temps de porte (1, 2, 3, …)
                TypeMesure.Interval     => "inter",
                TypeMesure.TachyOptique => "topti",
                TypeMesure.TachyContact => "tcont",
                TypeMesure.Stroboscope  => "strob",
                _ => "Mesure"
            };

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
        /// Écrit une valeur dans une cellule (<c>Range["A1"].Value2</c>) : en-têtes et labels de
        /// <see cref="AjouterFeuilleMesureAsync"/>.
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
        /// Écrit une formule dans la cellule d'une zone nommée portée feuille (équivalent de
        /// <see cref="EcrireValeurZoneNommeeInterne"/> pour <c>RefersToRange.Formula</c>). Excel
        /// rejette parfois <c>.Formula</c> (lien <c>[1]!</c> non résolu, syntaxe) alors que
        /// <c>.Formula2</c>/<c>.FormulaLocal</c> passent : on tente les variantes en cascade et
        /// on journalise chaque échec pour éviter un échec silencieux.
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
                                // Essai 1 : .Formula (US, virgules).
                                try
                                {
                                    refRange.Formula = formule;
                                    succes = true;
                                }
                                catch (Exception ex1)
                                {
                                    echecsAccumules += $".Formula KO: {ex1.Message}; ";
                                    // Essai 2 : .Formula2 (Excel 365+).
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
                        // Repli : Name portée classeur
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
        /// Calcule la variance via la macro VBA <c>Cal_variance</c> du Metrologo.xla
        /// (<c>Application.Run</c>) et l'écrit comme valeur dans <c>ZNVariance</c> : contournement
        /// du rejet 0x800A03EC de <c>=[1]!Cal_variance(...)</c> par <c>.Formula</c> en COM dynamic.
        /// Les arguments reproduisent la formule du rapport (SUMSQ des deltas de fréquence réelle,
        /// nombre de mesures, moyenne). Si <c>Run</c> échoue : repli C# <c>sumSq / (2*(N-1))</c>
        /// (variance d'Allan). Actuellement non appelée.
        /// </summary>
        private void CalculerEtEcrireVarianceViaComInterne(Mesure mesure, List<double> resultats,
            string nomFeuille)
        {
            if (_excel == null || _classeurActif == null) return;
            if (resultats == null || resultats.Count < 2) return;

            try
            {
                // 1. Fréquence réelle en C#, comme la formule
                //    `=IF(ISBLANK(ZNCoeffMult),E,(((E-1e7)/(POWER(10,Mult)*1e7))+1)*FNominale)`.
                var corriges = new List<double>(resultats.Count);
                foreach (var v in resultats) corriges.Add(ConvertirEnFreqReelleLocal(v, mesure));

                // 2. SUMSQ des carrés des différences consécutives (F_i - F_{i+1}).
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

                // 3. Macro VBA Cal_variance via Application.Run (repli C# pur si indisponible).
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
                    // Repli : variance d'Allan σ²(τ) = sumSq / (2 * (N-1))
                    if (nbMesures >= 2)
                    {
                        variance = sumSq / (2.0 * (nbMesures - 1));
                        JournalLog.Warn(CategorieLog.Excel, "VARIANCE_VIA_RUN_FALLBACK",
                            $"Application.Run KO ({exRun.Message}) — fallback C# Allan : "
                            + $"variance = sumSq/(2*(N-1)) = {variance:G6}");
                    }
                }

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
        /// Résout les coefficients d'incertitude du module sélectionné et les écrit dans
        /// <c>ZNCoeffA/B</c> (<c>C/D</c> pour les tachymètres) via COM, comme
        /// <c>ExcelService.EcrireStatsAsync</c>.
        /// </summary>
        private void EcrireCoeffsIncertitudeViaComInterne(Mesure mesure, List<double> resultats,
            string nomFeuille, double tempsGateSecondes)
        {
            // Stabilité : sigma relatif (n-1) mémorisé par gate pour calibrer l'axe Y log du graphe
            // à la dernière gate (comme ExcelService.EcrireStatsAsync, C# pur, sans écriture Excel).
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
        /// Redirige les liens externes vers un <c>Metrologo.xla</c> du classeur actif vers
        /// <see cref="Preferences.CheminMacroXLA"/> (<c>Workbook.ChangeLink</c>) : Excel nettoie le
        /// rels et résout aussitôt les formules <c>[1]!Cal_xxx(...)</c>. Remplace le
        /// PatcherLienMacroXLA (ClosedXML), non appelé en mode visible.
        /// </summary>
        private void RedirigerLinksXlaViaComInterne()
        {
            if (_excel == null || _classeurActif == null) return;
            string xlaPath = Preferences.CheminMacroXLA;
            if (string.IsNullOrWhiteSpace(xlaPath)) return;
            try
            {
                // LinkSources(xlLinkTypeExcelLinks=1) : tableau 1-based des chemins liés.
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
        /// Écrit une valeur dans une zone nommée portée classeur (zones partagées entre feuilles,
        /// ex : ZNNoFicheRecapFreq pour l'en-tête de la Récap).
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

        /// <summary>
        /// Plages nommées du récap de freq.xlsx (fréquence, tachy, strobo, intervalle : même mise
        /// en page), lues par le logiciel externe. Pas de plages en stabilité : son récap est encore
        /// exploité à la main.
        /// </summary>
        private static readonly (string Nom, string Colonne)[] _plagesRecap =
        {
            ("Recap_FreqNominale",  "A"),
            ("Recap_TempsMesure",   "B"),   // libellé de la gate, ex. « 1 s »
            ("Recap_FreqCorrigee",  "C"),
            ("Recap_IncertGlobale", "I"),
        };

        /// <summary>
        /// (Re)pose les plages nommées du récap sur les lignes remplies (de la ligne suivant la
        /// ligne modèle ZNRecapF_Ligne0, ligne 6, à la dernière dont la colonne A est non vide).
        /// Appelé à chaque ligne ajoutée et à la réinitialisation ; récap vide : la plage se
        /// réduit à sa 1re ligne. À appeler sous lock <c>_sync</c>.
        /// </summary>
        private void MettreAJourPlagesNommeesRecapInterne(dynamic recap)
        {
            if (_classeurActif == null) return;
            try
            {
                int ligneModele = 6;
                try
                {
                    dynamic ligne0 = _classeurActif.Names.Item("ZNRecapF_Ligne0");
                    try { ligneModele = (int)ligne0.RefersToRange.Row; }
                    finally { try { Marshal.ReleaseComObject(ligne0); } catch { } }
                }
                catch { /* zone absente ou cassée : valeurs du template */ }

                bool ColonneARemplie(int r)
                {
                    dynamic? cell = null;
                    try
                    {
                        cell = recap.Cells[r, 1];
                        try { return (Convert.ToString(cell.Formula) ?? "").Length > 0; } catch { return false; }
                    }
                    catch { return false; }
                    finally { if (cell != null) { try { Marshal.ReleaseComObject(cell); } catch { } } }
                }

                // 1re ligne de mesure : d'habitude celle sous la ligne modèle, mais la ligne modèle
                // elle-même accueille la 1re mesure après un « Relancer → écraser » (récap vidé).
                int premiere = ligneModele + 1;
                for (int r = ligneModele; r <= ligneModele + 2; r++)
                    if (ColonneARemplie(r)) { premiere = r; break; }

                int derniere = premiere;
                for (int r = premiere; r <= 300 && ColonneARemplie(r); r++)
                    derniere = r;

                dynamic noms = _classeurActif.Names;
                try
                {
                    foreach (var (nom, col) in _plagesRecap)
                    {
                        try { noms.Add(nom, $"='Récap.'!${col}${premiere}:${col}${derniere}"); }
                        catch (Exception ex)
                        {
                            JournalLog.Warn(CategorieLog.Excel, "PLAGE_RECAP_KO",
                                $"Plage nommée {nom} non posée : {ex.Message}");
                        }
                    }
                }
                finally { try { Marshal.ReleaseComObject(noms); } catch { } }
            }
            catch { /* best-effort : l'absence des plages ne doit pas casser la mesure */ }
        }

        /// <summary>
        /// Repli si l'enregistrement échoue (typiquement : rapport sur le réseau, M: perdu en cours
        /// de mesure). Enregistre sous le même nom dans le dossier local de la FI (Bureau\Metrologo\FI),
        /// bascule la FI en local pour la session et l'inscrit au rapatriement automatique.
        /// À appeler sous lock <c>_sync</c>.
        /// </summary>
        private void EnregistrerEnLocalApresEchecInterne(Mesure mesure)
        {
            if (_classeurActif == null || string.IsNullOrWhiteSpace(mesure.NumFI)) return;
            try
            {
                string dossierLocal = CheminsMetrologo.BasculerFIEnLocal(mesure.NumFI);
                string nom = Path.GetFileName(string.IsNullOrEmpty(_cheminClasseurActif)
                    ? (mesure.TypeMesure == TypeMesure.Stabilite ? "stab.xlsx" : "freq.xlsx")
                    : _cheminClasseurActif);
                string cible = Path.Combine(dossierLocal, nom);
                if (string.Equals(Path.GetFullPath(cible), Path.GetFullPath(_cheminClasseurActif ?? ""),
                        StringComparison.OrdinalIgnoreCase))
                    return; // déjà en local : rien de plus à tenter

                try { _excel!.DisplayAlerts = false; } catch { }
                _classeurActif.SaveAs(cible);
                _cheminClasseurActif = cible;
                try { _fullNameActifExcel = (string)_classeurActif.FullName; } catch { _fullNameActifExcel = cible; }
                TransfertReseauService.SignalerFIEnLocal(mesure.NumFI);
                JournalLog.Warn(CategorieLog.Excel, "EXCEL_SAVE_REPLI_LOCAL",
                    $"Rapport enregistré en local après échec de l'enregistrement : {cible} "
                    + "(rapatrié sur le réseau dès son retour).");
            }
            catch (Exception ex)
            {
                JournalLog.Erreur(CategorieLog.Excel, "EXCEL_SAVE_REPLI_LOCAL_KO",
                    $"Enregistrement de secours en local impossible : {ex.Message}");
            }
        }

        /// <summary>Écrit une valeur dans une zone nommée portée feuille (repli sur la portée classeur).</summary>
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
        /// Re-protège toutes les feuilles du classeur actif avec <see cref="MotDePasseProtectionFeuille"/>,
        /// en fin de mesure juste avant le Save : le rapport n'est plus modifiable dans Excel sans
        /// le mot de passe, que l'app connaît et utilise pour déprotéger à chaque relance / ajout de
        /// mesure. Chaque feuille est d'abord déprotégée (Protect lève si déjà protégée).
        /// Best-effort par feuille.
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
                try { recap.Unprotect("METROL"); } catch { try { recap.Unprotect("metrol"); } catch { try { recap.Unprotect(); } catch { } } }

                // Colonne D (date) élargie pour éviter « #### ».
                try { recap.Columns["D"].ColumnWidth = 13.0; } catch { }

                // En-tête fiche : N° FI + date du jour (écrase les valeurs héritées du template).
                EcrireValeurZoneNommeeClasseurInterne("ZNNoFicheRecapFreq", mesure.NumFI);
                EcrireValeurZoneNommeeClasseurInterne("ZNDateRecapFreq",
                    DateTime.Now.ToString("dd/MM/yyyy"));

                // 1re ligne vide après l'entête (ligne 5)
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
                // Col 5 : fréquence indiquée. Générateur : « Géné. ». Fréquencemètre : cellule laissée
                // vide, la valeur saisie en post-mesure y est écrite par EcrireFreqIndiqueeRecapAsync.
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
                // Col 11 : n°Module, lu dans la cellule « Module = » de la feuille (F10, G10 en tachy).
                string celModuleRecap = EnTetesMesureHelper.EstTachymetre(mesure.TypeMesure) ? "G10" : "F10";
                EcrireFormuleRecapInterne(recap, nouvelleLigne, 11, $"={qf}{celModuleRecap}");

                // Écritures scientifiques (écart-type + incertitudes) : forcées à 2 décimales.
                FormaterScientifiqueRecapInterne(recap, nouvelleLigne, 4, 6, 7, 8, 9);

                // En-tête K si vide
                EcrireValeurRecapSiVideInterne(recap, ligneEntete, 11, "n°Module");

                MettreAJourPlagesNommeesRecapInterne(recap);
            }
            finally
            {
                try { Marshal.ReleaseComObject(recap); } catch { }
            }
        }

        /// <summary>
        /// Écrit la « fréquence indiquée » saisie en post-mesure (mode Fréquencemètre) dans la
        /// colonne 5 de la ligne Récap de la feuille active, là où le mode Générateur écrit « Géné. ».
        /// Simple report d'affichage : la feuille de mesure n'est pas touchée et ZNFreqRef /
        /// ZNFreqCorr ne sont pas alimentés. La ligne est repérée par la formule de sa colonne 1
        /// (<c>='&lt;feuille&gt;'!ZNFreqMoyReel</c>).
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

                try { recap.Columns["D"].ColumnWidth = 13.0; } catch { }

                // En-tête fiche
                EcrireValeurRecapInterne(recap, 1, 2, mesure.NumFI);
                EcrireValeurRecapInterne(recap, 1, 4, DateTime.Now.ToString("dd/MM/yyyy"));

                // N° de gate lu dans le nom de feuille (repli : compteur de lignes remplies)
                int numero;
                int i = nomFeuille.Length;
                while (i > 0 && char.IsDigit(nomFeuille[i - 1])) i--;
                string chiffres = nomFeuille.Substring(i);
                if (!int.TryParse(chiffres, out numero))
                {
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
                // Valeurs Mini : formule d'origine de Stab1.xls, plancher 1E-12 quand sigma ≤
                // incertitude (« * » en col I). Un 0 rendait le point invisible sur l'axe log.
                EcrireFormuleRecapInterne(recap, ligne, 8,
                    $"=IF(ISBLANK(C{ligne}),,IF((C{ligne}-F{ligne})<=0,0.000000000001,MAX(0.000000000001,C{ligne}-F{ligne})))");
                // n°Module lu dans la cellule « Module = » (F10).
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
        /// Force le format scientifique à 2 décimales sur des cellules d'une ligne du récap
        /// (incertitudes, écart-type). <c>NumberFormatLocal</c> avec la VIRGULE (« 0,00E+00 ») :
        /// sur un Excel français, un format à point posé via <c>NumberFormat</c> est mal interprété
        /// (mantisse non normalisée). Repli sur le format US, best-effort par cellule.
        /// </summary>
        private static void FormaterScientifiqueRecapInterne(dynamic recap, int row, params int[] colonnes)
        {
            foreach (int col in colonnes)
            {
                dynamic? c = null;
                try
                {
                    c = recap.Cells[row, col];
                    try { c.NumberFormatLocal = "0,00E+00"; }
                    catch { try { c.NumberFormat = "0.00E+00"; } catch { } }
                }
                catch { /* best-effort */ }
                finally { if (c != null) { try { Marshal.ReleaseComObject(c); } catch { } } }
            }
        }

        private static void EcrireFormuleRecapInterne(dynamic recap, int row, int col, string formule)
        {
            try
            {
                dynamic c = recap.Cells[row, col];
                try
                {
                    // Le template METROLOGO.xlsx a une ligne 6 pré-stylée dont les cellules sont typées
                    // t="s" (shared string) : y écrire .Formula stocke alors la formule COMME TEXTE et
                    // Excel affiche « ='freq1'!ZNLibGate » littéralement. NumberFormat = "General" et
                    // Clear() ne changent pas ce type ; seule l'écriture d'un nombre AVANT la formule
                    // le convertit en t="n".
                    try { c.Clear(); } catch { }
                    try { c.Style = "Normal"; } catch { }
                    try { c.NumberFormat = "General"; } catch { }
                    try { c.Value2 = 0.0; } catch { }   // force le type numérique (t="s" → t="n")
                    c.Formula = formule;

                    // Centrage (Clear() a effacé l'alignement hérité du template) ; xlCenter = -4108.
                    try { c.HorizontalAlignment = -4108; } catch { }
                    try { c.VerticalAlignment = -4108; } catch { }

                    // Ceinture et bretelles : si la cellule est quand même restée en texte, on passe par
                    // Application.Evaluate (affichage correct, mais plus de recalcul live).
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
        /// Ajuste les plages des séries du graphe Stab au nombre réel de gates balayées (équivalent
        /// COM de <c>ExcelService.RendreGrapheStabDynamique</c>, via le modèle objet des graphiques
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
                                    // Plages patchées par regex sur la formule de la série.
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
                            // Axe Y en auto (retire les bornes figées 1E-12/1E-08 du Stab1.xls)
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
        /// Calibre l'axe Y log du graphe Stab (après recalcul) sur les lignes 6..5+nbGates de la
        /// Récap. : écart type (C), Valeurs Maxi (G) et Valeurs Mini réelles (C − F si > 0) ; repli
        /// sur les sigmas si la Récap ne donne rien. Place aussi l'axe X en bas (croisement au
        /// minimum), le croisement fixe à 1E-12 de Stab1.xls tombant au milieu du graphe sous 1E-12.
        /// <para>Plancher des Valeurs Mini : dans Stab1.xls un Mini ≤ 0 (sigma ≤ incertitude) vaut
        /// 1E-12, aussi bas du graphe (« hors échelle, en bas »). Si l'échelle descend plus bas
        /// (sigmas &lt; 1E-12 aux longs temps de porte), ce plancher passerait AU-DESSUS de l'écart
        /// type : il suit donc min(1E-12, borne basse).</para>
        /// </summary>
        private void CalibrerAxeYGrapheStabViaComInterne(int nbGates, IEnumerable<double> sigmas)
        {
            if (_classeurActif == null) return;
            var lignes = LireLignesRecapStab(nbGates);
            var positifs = new List<double>();
            foreach (var (_, c, f, g) in lignes)
            {
                if (c > 0) positifs.Add(c);
                if (g > 0) positifs.Add(g);
                if (c - f > 0) positifs.Add(c - f);   // Mini réel ; le plancher n'entre pas dans le calcul
            }
            if (positifs.Count == 0) positifs = sigmas.Where(s => s > 0).ToList();
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
                    const double PLANCHER_STAB1 = 1e-12;
                    double plancher = Math.Min(PLANCHER_STAB1, minBorne);
                    if (plancher < PLANCHER_STAB1)
                    {
                        string pl = plancher.ToString("0.###E+0", System.Globalization.CultureInfo.InvariantCulture);
                        foreach (var (ligne, _, _, _) in lignes)
                            EcrireFormuleRecapInterne(ws, ligne, 8,
                                $"=IF(ISBLANK(C{ligne}),,IF((C{ligne}-F{ligne})<=0,{pl},MAX({pl},C{ligne}-F{ligne})))");
                        try { ws.Calculate(); } catch { }
                        JournalLog.Info(CategorieLog.Excel, "GRAPHE_STAB_PLANCHER_MINI",
                            $"Plancher des Valeurs Mini abaissé à {pl} (bas de l'échelle, sigmas < 1E-12).");
                    }

                    dynamic chartObjects = ws.ChartObjects();
                    int nbCharts = chartObjects.Count;
                    for (int i = 1; i <= nbCharts; i++)
                    {
                        dynamic chartObj = chartObjects.Item(i);
                        try
                        {
                            // Le template ancre le graphe en ligne 18 (prévu pour 10 temps de porte) : à 13
                            // temps il recouvrait la ligne 100 s. On le place sous la dernière ligne + 1 vide.
                            try
                            {
                                dynamic ancre = ws.Range($"A{5 + nbGates + 2}");
                                chartObj.Top = (double)ancre.Top;
                                try { Marshal.ReleaseComObject(ancre); } catch { }
                            }
                            catch (Exception ex)
                            {
                                JournalLog.Warn(CategorieLog.Excel, "GRAPHE_STAB_POSITION_KO",
                                    $"Déplacement du graphe Stab échoué : {ex.Message}");
                            }

                            dynamic chart = chartObj.Chart;
                            try
                            {
                                dynamic valAxis = chart.Axes(2);   // xlValue
                                valAxis.MinimumScale = minBorne;
                                valAxis.MaximumScale = maxBorne;
                                valAxis.Crosses = 4;               // xlAxisCrossesMinimum : axe X en bas
                                try { Marshal.ReleaseComObject(valAxis); } catch { }

                                // Traits Maxi → Mini (HiLoLines) de Stab1.xls : Excel les retire dès le 1er
                                // enregistrement COM (gate 1, Récap vide) ; on les repose avant le Save final.
                                dynamic groupe = chart.ChartGroups(1);
                                groupe.HasHiLoLines = true;
                                dynamic bordure = groupe.HiLoLines.Border;
                                bordure.LineStyle = 1;             // xlContinuous
                                bordure.Weight = 1;                // xlHairline (0,25 pt comme le template)
                                bordure.Color = 0;                 // noir
                                try { Marshal.ReleaseComObject(bordure); } catch { }
                                try { Marshal.ReleaseComObject(groupe); } catch { }
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

        /// <summary>Lignes calculées de la Récap. Stab (6..5+nbGates) : écart type (C), incertitude
        /// globale (F), Valeurs Maxi (G). Seules les lignes dont l'écart type est un nombre sont rendues.</summary>
        private List<(int Ligne, double C, double F, double G)> LireLignesRecapStab(int nbGates)
        {
            var lignes = new List<(int, double, double, double)>();
            if (_classeurActif == null || nbGates <= 0) return lignes;
            try
            {
                dynamic ws = _classeurActif.Worksheets.Item("Récap.");
                try
                {
                    double Lire(string adresse)
                    {
                        dynamic cell = ws.Range(adresse);
                        try { return cell.Value2 is double d ? d : double.NaN; }
                        finally { try { Marshal.ReleaseComObject(cell); } catch { } }
                    }
                    for (int ligne = 6; ligne <= 5 + nbGates; ligne++)
                    {
                        double c = Lire($"C{ligne}");
                        if (double.IsNaN(c)) continue;
                        double f = Lire($"F{ligne}");
                        lignes.Add((ligne, c, double.IsNaN(f) ? 0 : f, Lire($"G{ligne}")));
                    }
                }
                finally { try { Marshal.ReleaseComObject(ws); } catch { } }
            }
            catch { /* best-effort : repli sur les sigmas */ }
            return lignes;
        }

        /// <summary>
        /// Amène la fenêtre du rapport actif au premier plan (Activate, xlMaximized, BringWindowToTop,
        /// SetForegroundWindow) : annule la « fenêtre grise par-dessus » d'Excel SDI quand une
        /// opération COM (Copy, Save, Open) met une shell en TopMost. SetForegroundWindow échoue
        /// silencieusement si notre thread n'a pas le focus : on l'attache temporairement au thread
        /// foreground via AttachThreadInput.
        /// </summary>
        private void ForcerRapportAuPremierPlanInterne()
        {
            if (_excel == null || _classeurActif == null) return;
            try
            {
                // 1. Activate côté Excel (Z-order Excel uniquement, pas celui de Windows).
                try { _classeurActif.Activate(); } catch { }

                // 2. xlMaximized : le rapport masque visuellement toute autre fenêtre Excel dessous.
                try
                {
                    dynamic w = _classeurActif.Windows[1];
                    try { w.WindowState = -4137; /* xlMaximized */ }
                    finally { try { Marshal.ReleaseComObject(w); } catch { } }
                }
                catch { }

                // 3. Win32 : le rapport passe devant toute autre fenêtre (y compris une shell TopMost),
                //    avec l'astuce AttachThreadInput (cf. summary).
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
        /// Démarre le timer (500 ms) qui appelle <see cref="MasquerFenetresShellExcelInterne"/> tant
        /// qu'un classeur est ouvert : aucune shell SDI ne reste visible plus d'une demi-seconde,
        /// même si l'appel initial a raté (race COM). Le tick est léger (1 appel COM + EnumWindows).
        /// Ne PAS y appeler <see cref="FermerClasseursParasitesInterne"/> : il itère sur les Workbooks
        /// (~100-300 ms) sous le lock <c>_sync</c> partagé avec <see cref="EcrireValeurLiveAsync"/> et
        /// gèlerait l'affichage live des mesures du compteur GPIB. Ce nettoyage est déjà fait dans
        /// <c>OuvrirEtAfficherAsync</c> et <c>AjouterFeuilleMesureAsync</c>, seuls moments où une
        /// shell peut apparaître.
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
                            // 1 appel COM (Hwnd) + EnumWindows : ~1-2 ms, négligeable pour les écritures live.
                            MasquerFenetresShellExcelInterne();
                        }
                    }
                    catch { /* best-effort, ne doit jamais lever depuis un Timer */ }
                }, null, 500, 500);
            }
        }

        /// <summary>
        /// Arrête le timer de masquage (fermeture du classeur, Dispose). Idempotent.
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
        /// Masque les fenêtres top-level d'Excel autres que le rapport actif (shell SDI grise,
        /// Book1 résiduel…). Elles ne figurent pas dans <c>Application.Windows</c> (Workbook Windows
        /// seulement) d'où l'EnumWindows Win32. Le critère est le HWND du rapport, pas le titre,
        /// qui varie selon la langue/version d'Excel.
        /// </summary>
        private void MasquerFenetresShellExcelInterne()
        {
            if (_excel == null || _excelPid <= 0 || _classeurActif == null) return;
            try
            {
                // HWND du rapport actif : seule fenêtre top-level du process gardée visible.
                IntPtr hwndRapport = IntPtr.Zero;
                try
                {
                    dynamic w = _classeurActif.Windows[1];
                    try { hwndRapport = new IntPtr((int)w.Hwnd); }
                    finally { try { Marshal.ReleaseComObject(w); } catch { } }
                }
                catch { /* best-effort */ }

                // Garde-fou critique : sans HWND du rapport (race COM, fenêtre pas encore créée) on
                // ne masque RIEN, sinon le rapport lui-même serait caché.
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

                        // Excel SDI peut ré-afficher la shell sur certaines opérations COM (SaveAs,
                        // Workbooks.Open...) : on cumule les parades.
                        // 1. SW_HIDE ; 2. SW_FORCEMINIMIZE (si réaffichée : minimisée) ;
                        // 3. HWND_BOTTOM (sous toute fenêtre visible) ; 4. hors écran (-32000, -32000).
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

        /// <summary>Vrai si ce classeur (nom complet donné par Excel ou la ROT) est notre rapport
        /// actif : même nom complet qu'à l'ouverture, ou même chemin normalisé. Évite de prendre
        /// notre rapport pour un « parasite » à masquer/fermer.</summary>
        private bool EstClasseurActif(string fullName)
        {
            if (string.IsNullOrEmpty(fullName)) return false;
            if (!string.IsNullOrEmpty(_fullNameActifExcel)
                && string.Equals(fullName, _fullNameActifExcel, StringComparison.OrdinalIgnoreCase))
                return true;
            return ChemPathEgal(fullName, _cheminClasseurActif)
                || ChemPathEgal(fullName, _fullNameActifExcel);
        }

        /// <summary>Compare deux chemins après <see cref="NormaliserChemin"/> (casse, 8.3, lecteur réseau/UNC).</summary>
        private static bool ChemPathEgal(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return false;
            try
            {
                return string.Equals(NormaliserChemin(a), NormaliserChemin(b), StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
            }
        }

        /// <summary>Vrai si <paramref name="chemin"/> est dans <paramref name="dossier"/> (lecteur
        /// réseau et chemin UNC considérés équivalents).</summary>
        private static bool EstSousDossier(string chemin, string dossier)
        {
            if (string.IsNullOrWhiteSpace(chemin) || string.IsNullOrWhiteSpace(dossier)) return false;
            try
            {
                string c = NormaliserChemin(chemin);
                string d = NormaliserChemin(dossier).TrimEnd('\\') + "\\";
                return c.StartsWith(d, StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        private static readonly Dictionary<string, string?> _racinesUnc = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Chemin complet, la lettre d'un lecteur RÉSEAU étant remplacée par sa racine UNC
        /// (<c>M:\x</c> → <c>\\serveur\partage\x</c>). Excel désigne un fichier réseau par l'une ou
        /// l'autre forme : sans cela, notre rapport ouvert sur M: n'était pas reconnu.
        /// </summary>
        private static string NormaliserChemin(string chemin)
        {
            string complet = Path.GetFullPath(chemin);
            if (complet.Length >= 2 && complet[1] == ':')
            {
                string lecteur = complet.Substring(0, 2);
                string? unc;
                lock (_racinesUnc)
                {
                    if (!_racinesUnc.TryGetValue(lecteur, out unc))
                    {
                        unc = null;
                        try
                        {
                            var sb = new System.Text.StringBuilder(512);
                            int taille = sb.Capacity;
                            if (WNetGetConnection(lecteur, sb, ref taille) == 0)
                                unc = sb.ToString().TrimEnd('\\');
                        }
                        catch { /* pas un lecteur réseau */ }
                        _racinesUnc[lecteur] = unc;
                    }
                }
                if (!string.IsNullOrEmpty(unc))
                    complet = unc + complet.Substring(2);
            }
            return complet.TrimEnd('\\');
        }

        [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
        private static extern int WNetGetConnection(string localName, System.Text.StringBuilder remoteName, ref int length);

        private void FermerClasseurActifInterne()
        {
            ArreterTimerMasquageShellInterne();

            // Cache l'instance AVANT de fermer le classeur : visible sans classeur, Excel afficherait
            // la shell SDI grise. OuvrirEtAfficherAsync remettra Visible=true après l'ouverture suivante.
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
            _fullNameActifExcel = string.Empty;

            // Libère aussitôt les handles fichier d'Excel (RCW) : sans cela le fichier reste verrouillé
            // 100-500 ms après le Close et ClosedXML bascule sur un nom horodaté à la mesure suivante.
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Quitte l'instance Excel hôte, à la fermeture de l'application.
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
