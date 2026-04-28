using Metrologo.Services.Journal;
using System;
using System.Collections.Generic;
using System.IO;
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

        private readonly object _sync = new();
        private bool _disposed;

        private ExcelInteropHost() { }

        /// <summary>Vrai si Excel est démarré et prêt à ouvrir un classeur.</summary>
        public bool EstDemarre => _excel != null;

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

                    _excel = Activator.CreateInstance(excelType);
                    if (_excel == null) return;

                    _excel.Visible = false;
                    _excel.DisplayAlerts = false;
                    _excel.ScreenUpdating = false;
                    _excel.AskToUpdateLinks = false;

                    JournalLog.Info(CategorieLog.Excel, "EXCEL_HOTE_DEMARRE",
                        "Instance Excel cachée prête en arrière-plan.");
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
                    if (_excel == null) return;

                    // Ferme tout classeur précédent (une seule mesure live à la fois).
                    FermerClasseurActifInterne();

                    // Workbooks.Open(Filename, UpdateLinks, ReadOnly, …). On passe seulement
                    // les 2 premiers paramètres — UpdateLinks=0 = ne pas mettre à jour les liens
                    // externes au chargement (évite le dialogue bloquant sur le .xla absent).
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

                    _excel.ScreenUpdating = true;
                    _excel.Visible = true;

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
                    try
                    {
                        int row = ligneDebut + indexMesure;
                        // Cells(row, col) : col=1 (A) = HEURE, col=2 (B) = mesure.
                        _feuilleMesure.Cells[row, 1].Value2 = horodatage.ToString("HH:mm:ss");
                        _feuilleMesure.Cells[row, 2].Value2 = valeur;
                    }
                    catch (Exception ex)
                    {
                        // Écriture live best-effort : une cellule plantée ne doit pas tuer la mesure.
                        JournalLog.Warn(CategorieLog.Excel, "EXCEL_ECRITURE_LIVE_ERREUR",
                            $"Écriture live impossible (i={indexMesure}) : {ex.Message}");
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

                        // Range A{ligneDebut}:B{ligneDebut+N-1}
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

        /// <summary>
        /// Ajoute un graphe de stabilité (XY Scatter) sur la feuille <c>Récap.</c> du classeur
        /// actif, reproduisant la structure historique Stab1.xls : 3 séries (Écart type / Valeurs
        /// Maxi / Valeurs Mini) en fonction du Temps de porte (col A). Idempotent : si un graphe
        /// nommé <c>GrapheStab</c> existe déjà, ne fait rien.
        /// </summary>
        /// <param name="nomFeuilleRecap">Nom de la feuille Récap. (par convention « Récap. »).</param>
        public Task AjouterGrapheStabiliteAsync(string nomFeuilleRecap = "Récap.")
            => Task.Run(() =>
            {
                lock (_sync)
                {
                    if (_classeurActif == null) return;
                    try
                    {
                        // Localise la feuille Récap.
                        dynamic? recap = null;
                        foreach (dynamic ws in _classeurActif.Worksheets)
                        {
                            if (string.Equals((string)ws.Name, nomFeuilleRecap, StringComparison.OrdinalIgnoreCase))
                            {
                                recap = ws;
                                break;
                            }
                        }
                        if (recap == null) return;

                        // Idempotence : si un graphe « GrapheStab » est déjà présent, on n'en
                        // recrée pas un (utile quand l'utilisateur relance un balayage sur le même FI).
                        dynamic chartObjects = recap.ChartObjects();
                        for (int i = 1; i <= (int)chartObjects.Count; i++)
                        {
                            dynamic existing = chartObjects.Item(i);
                            if (string.Equals((string)existing.Name, "GrapheStab", StringComparison.OrdinalIgnoreCase))
                            {
                                Marshal.ReleaseComObject(existing);
                                Marshal.ReleaseComObject(chartObjects);
                                Marshal.ReleaseComObject(recap);
                                return;
                            }
                            Marshal.ReleaseComObject(existing);
                        }
                        Marshal.ReleaseComObject(chartObjects);

                        // Position calquée sur Stab1.xls historique : sous la zone des données,
                        // pleine largeur. L'utilisateur peut le déplacer/redimensionner.
                        dynamic co = recap.ChartObjects().Add(0, 217, 509, 376); // L, T, W, H (pts)
                        co.Name = "GrapheStab";
                        dynamic chart = co.Chart;

                        // xlXYScatter (65) : nuage de points seuls (sans lignes), comme dans
                        // Stab1.xls. Pour un graphe de stabilité (Allan deviation), les points
                        // sont plus lisibles que des lignes connectées.
                        chart.ChartType = 65;

                        chart.HasTitle = true;
                        chart.ChartTitle.Text = "STABILITE";

                        // Axe Y log + minor gridlines (aspect quadrillé typique des graphes Allan
                        // deviation) ; titres d'axes et légende en bas alignés sur Stab1.xls.
                        try
                        {
                            dynamic axeY = chart.Axes(2);  // 2 = xlValue
                            axeY.ScaleType = -4133;            // xlScaleLogarithmic
                            axeY.HasMajorGridlines = false;
                            axeY.HasMinorGridlines = true;
                            axeY.HasTitle = true;
                            axeY.AxisTitle.Text = "Stabilité relative";
                            Marshal.ReleaseComObject(axeY);

                            dynamic axeX = chart.Axes(1);  // 1 = xlCategory
                            axeX.HasMajorGridlines = false;
                            axeX.HasMinorGridlines = false;
                            axeX.HasTitle = true;
                            axeX.AxisTitle.Text = "Temps de Mesure (s)";
                            Marshal.ReleaseComObject(axeX);

                            chart.HasLegend = true;
                            chart.Legend.Position = -4107;     // xlLegendPositionBottom
                        }
                        catch { /* config axes/légende non bloquante */ }

                        // 3 séries avec markers distincts pour les différencier visuellement,
                        // alignés sur la convention Stab1.xls : Square / X / Dash.
                        var sc = chart.SeriesCollection();

                        dynamic s1 = sc.NewSeries();
                        s1.Name = $"='{nomFeuilleRecap}'!$C$3";   // En-tête « Ecart type »
                        s1.XValues = $"='{nomFeuilleRecap}'!$A$6:$A$15";
                        s1.Values  = $"='{nomFeuilleRecap}'!$C$6:$C$15";
                        try { s1.MarkerStyle = 2; s1.MarkerSize = 5; } catch { }   // xlMarkerStyleSquare

                        dynamic s2 = sc.NewSeries();
                        s2.Name = $"='{nomFeuilleRecap}'!$G$3";   // « Valeurs Maxi. »
                        s2.XValues = $"='{nomFeuilleRecap}'!$A$6:$A$15";
                        s2.Values  = $"='{nomFeuilleRecap}'!$G$6:$G$15";
                        try { s2.MarkerStyle = -4115; s2.MarkerSize = 5; } catch { } // xlMarkerStyleX

                        dynamic s3 = sc.NewSeries();
                        s3.Name = $"='{nomFeuilleRecap}'!$H$3";   // « Valeurs Mini. »
                        s3.XValues = $"='{nomFeuilleRecap}'!$A$6:$A$15";
                        s3.Values  = $"='{nomFeuilleRecap}'!$H$6:$H$15";
                        try { s3.MarkerStyle = -4118; s3.MarkerSize = 5; } catch { } // xlMarkerStyleDash

                        // High-Low Lines : barres verticales reliant les 3 séries (Maxi / Mini /
                        // Écart type) à chaque temps de porte — c'est ce qui donne l'aspect
                        // « error bars » propre du Stab1.xls historique. Sans ça, on n'a que
                        // 3 points isolés. La propriété est sur le ChartGroup, pas sur la série.
                        try
                        {
                            dynamic groupe = chart.ChartGroups(1);
                            groupe.HasHiLoLines = true;
                            Marshal.ReleaseComObject(groupe);
                        }
                        catch { /* propriété non dispo selon firmware Excel — non bloquant */ }

                        Marshal.ReleaseComObject(s1);
                        Marshal.ReleaseComObject(s2);
                        Marshal.ReleaseComObject(s3);
                        Marshal.ReleaseComObject(sc);
                        Marshal.ReleaseComObject(chart);
                        Marshal.ReleaseComObject(co);
                        Marshal.ReleaseComObject(recap);

                        JournalLog.Info(CategorieLog.Excel, "EXCEL_GRAPHE_STAB",
                            "Graphe de stabilité ajouté dans la feuille Récap.");
                    }
                    catch (Exception ex)
                    {
                        JournalLog.Warn(CategorieLog.Excel, "EXCEL_GRAPHE_STAB_ERR",
                            $"Création du graphe de stabilité impossible : {ex.Message}");
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

        private void FermerClasseurActifInterne()
        {
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
