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
                        // Décalage de 3 colonnes : col=4 (D) = HEURE, col=5 (E) = mesure brute.
                        // Les colonnes 1-3 (A/B/C) sont réservées à n°Module/Fonction/Condition 1
                        // (écrites une fois par ExcelService.InitialiserRapportAsync).
                        _feuilleMesure.Cells[row, 4].Value2 = horodatage.ToString("HH:mm:ss");
                        _feuilleMesure.Cells[row, 5].Value2 = valeur;
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

            try
            {
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return;
                _excel = Activator.CreateInstance(excelType);
                if (_excel == null) return;
                _excel.Visible = false;
                _excel.DisplayAlerts = false;
                _excel.ScreenUpdating = false;
                _excel.AskToUpdateLinks = false;
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "EXCEL_HOTE_REDEMARRAGE_KO",
                    $"Redémarrage Excel échoué : {ex.Message}");
                _excel = null;
            }
        }

        /// <summary>
        /// Ferme tous les classeurs ouverts dans l'instance Excel hôte SAUF
        /// <see cref="_classeurActif"/>. Évite que des fenêtres grises parasites (Book1
        /// par défaut, anciens classeurs) restent visibles à côté du classeur de mesure.
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
                for (int i = nb; i >= 1; i--)
                {
                    dynamic wb = _excel.Workbooks[i];
                    try
                    {
                        // Compare via le chemin complet — comparer les références dynamic
                        // est peu fiable en COM (proxy différent à chaque accès).
                        string fullName = (string)wb.FullName;
                        if (!string.Equals(fullName, _cheminClasseurActif,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            try { wb.Close(false); }   // SaveChanges = false : pas de save sur classeurs parasites
                            catch { /* best-effort */ }
                        }
                    }
                    catch { /* best-effort */ }
                    finally
                    {
                        try { Marshal.ReleaseComObject(wb); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Excel, "EXCEL_PARASITES_KO",
                    $"Nettoyage des classeurs parasites échoué : {ex.Message}");
            }
        }

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

            // Re-cache l'instance Excel après fermeture du classeur — sinon une fenêtre
            // Excel grise sans contenu reste visible à l'utilisateur (l'instance hôte
            // est conservée pour les ouvertures rapides suivantes, mais doit redevenir
            // invisible tant qu'aucun classeur n'est dedans).
            if (_excel != null)
            {
                try { _excel.Visible = false; } catch { /* best-effort */ }
            }
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
