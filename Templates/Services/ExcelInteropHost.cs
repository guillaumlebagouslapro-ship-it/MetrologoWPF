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
