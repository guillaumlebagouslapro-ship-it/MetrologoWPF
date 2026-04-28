using ClosedXML.Excel;
using Metrologo.Models;
using Metrologo.Services.Catalogue;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Metrologo.Services
{
    public interface IExcelService
    {
        /// <summary>
        /// Initialise une feuille de mesure dans le classeur du FI. Le paramètre optionnel
        /// <paramref name="gateIndexOverride"/> permet de spécifier la gate à inscrire dans
        /// les zones nommées (<c>ZNGate</c>/<c>ZNLibGate</c>/<c>ZNValGateSecondes</c>) — utile
        /// pour les balayages de stabilité où chaque feuille est associée à une gate différente
        /// sans qu'on souhaite muter <c>config.GateIndex</c> partout.
        /// </summary>
        Task InitialiserRapportAsync(string numeroFI, Mesure configuration, Rubidium rubidium, int? gateIndexOverride = null);

        /// <summary>
        /// Pré-insère les lignes vides nécessaires pour les mesures à venir (au-delà des 2 lignes
        /// par défaut du template). Les formules <c>Fréq. Réelle</c> et <c>F(i)-F(i+1)</c> sont
        /// ajoutées. Les cellules HEURE et mesure restent vides — elles seront remplies en direct.
        /// </summary>
        Task PreparerLignesMesureAsync(int nbMesures);

        /// <summary>
        /// Sauvegarde le classeur ClosedXML sur le disque (sans ouvrir Excel) pour qu'il puisse
        /// être repris ensuite par <see cref="ExcelInteropHost"/> en écriture live.
        /// </summary>
        Task<string> SauvegarderSurDisqueAsync();

        /// <summary>Nom de la feuille créée pour cette mesure (ex: Freq1, Stab1).</summary>
        string NomFeuilleMesure { get; }

        /// <summary>
        /// Écrit la moyenne et la variance dans les zones nommées — à appeler après la boucle
        /// de mesures, pour que le Récap. cross-sheet fonctionne.
        /// </summary>
        Task EcrireStatsAsync(List<double> resultats);

        /// <summary>
        /// Écrit les N mesures (HEURE + valeur) directement dans la feuille de mesure courante
        /// via ClosedXML, **sans passer par Excel/Interop**. Utilisé en mode « invisible » pour
        /// la Stabilité : Excel n'est jamais ouvert pendant la mesure, tout se fait en mémoire,
        /// et l'utilisateur voit le fichier final apparaître à la toute fin du balayage.
        /// </summary>
        Task EcrireValeursBatchClosedXMLAsync(int ligneDebut, IList<(DateTime ts, double valeur)> mesures);

        Task MettreAJourRecapFreqAsync(Mesure mesure);
        Task MettreAJourRecapStabAsync(Mesure mesure);

        /// <summary>
        /// Re-ouvre depuis le disque un classeur déjà initialisé, après qu'Interop l'ait rempli
        /// avec les valeurs live. Ne crée pas de nouvelle feuille : on récupère la feuille dont
        /// le nom est <see cref="NomFeuilleMesure"/>.
        /// </summary>
        Task RouvrirClasseurAsync();

        /// <summary>
        /// Sauvegarde finale après Recap + stats. Le fichier reste ouvert par <see cref="ExcelInteropHost"/>
        /// (l'utilisateur peut continuer à l'inspecter).
        /// </summary>
        Task SauvegarderFinalAsync();

        void FermerExcel();
    }

    public class ExcelService : IExcelService
    {
        // Nom de la feuille modèle — jamais modifiée.
        private const string NOM_MODELE = "ModFeuille";
        private const string NOM_RECAP = "Récap.";

        // Première ligne de mesures dans le template
        private const int LIGNE_DEBUT_MESURES = 9;

        // Zones nommées indiquant l'emplacement d'insertion des lignes Recap (cf. template xltm).
        private const string ZN_RECAPF_DEBZONE = "ZNRecapF_DebZone";
        private const string ZN_RECAPS_DEBZONE = "ZNRecapS_DebZone";

        // Lignes de fallback si les zones ZNRecapF/S_DebZone sont absentes (valeurs par défaut du template).
        private const int LIGNE_FALLBACK_RECAPF = 19;
        private const int LIGNE_FALLBACK_RECAPS = 19;

        private XLWorkbook? _workbook;
        private IXLWorksheet? _feuilleMesure;   // feuille NOUVELLEMENT créée pour cette mesure
        private string _cheminFichier = string.Empty;
        private string _nomFeuilleMesure = string.Empty;

        public string NomFeuilleMesure => _nomFeuilleMesure;

        /// <summary>
        /// Chemin du fichier Excel réellement utilisé pour cette mesure (peut différer de
        /// <c>Mesures_{FI}.xlsm</c> si un fallback timestampé a été appliqué — cf. Excel verrouillé).
        /// Exposé pour que la UI puisse en informer l'utilisateur.
        /// </summary>
        public string CheminFichierGenere => _cheminFichier;

        /// <summary>Vrai si le fichier a dû être écrit sous un nom de fallback au lieu du nom principal.</summary>
        public bool FallbackTimestampUtilise { get; private set; }

        public async Task InitialiserRapportAsync(string numeroFI, Mesure config, Rubidium rubidium, int? gateIndexOverride = null)
        {
            int gateInscrite = gateIndexOverride ?? config.GateIndex;
            bool estStab = config.TypeMesure == TypeMesure.Stabilite;

            await Task.Run(() =>
            {
                // --- 1. Détermination du dossier et du fichier ---
                //    La Stabilité utilise un fichier séparé (Récap. à 8 colonnes spécifique)
                //    pour ne pas polluer la Récap. Fréquence du fichier principal du FI.
                string numFISafe = SanitizerNomFichier(numeroFI);

                string dossier = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "Metrologo", numFISafe);
                Directory.CreateDirectory(dossier);

                string suffixe = estStab ? "_Stab" : string.Empty;
                _cheminFichier = Path.Combine(dossier, $"Mesures{suffixe}_{numFISafe}.xlsm");
                FallbackTimestampUtilise = false;

                // --- 2. Ouverture : fichier existant ou copie du template ---
                //    Template Stab dédié pour la Stabilité (Récap. 8 cols + zones nommées
                //    ZNRecapS_*), template Fréquence pour le reste.
                string nomTemplate = estStab ? "METROLOGO_Stab.xltm" : "METROLOGO.xltm";
                string templatePath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "Templates", nomTemplate);

                if (File.Exists(_cheminFichier))
                {
                    bool partirDuTemplate = false;

                    if (FichierEstVerrouille(_cheminFichier))
                    {
                        // Tentative 1 : fermer Excel poliment via WM_CLOSE (marche si pas de dialogue bloquant).
                        WindowHelper.FermerFenetresExcel(Path.GetFileName(_cheminFichier));
                        var sw = System.Diagnostics.Stopwatch.StartNew();
                        while (FichierEstVerrouille(_cheminFichier) && sw.ElapsedMilliseconds < 3000)
                        {
                            System.Threading.Thread.Sleep(150);
                        }

                        if (FichierEstVerrouille(_cheminFichier))
                        {
                            // Tentative 2 : fallback sur un nom de fichier timestampé — on n'interrompt
                            //                ni la mesure ni le travail de l'utilisateur dans Excel.
                            string cheminAlternatif = Path.Combine(
                                dossier,
                                $"Mesures_{numeroFI}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsm");

                            // On tente de copier le fichier existant pour garder Récap. + historique.
                            // Si Excel bloque aussi la lecture, on repart du template vierge.
                            try
                            {
                                File.Copy(_cheminFichier, cheminAlternatif, overwrite: false);
                            }
                            catch (IOException)
                            {
                                partirDuTemplate = true;
                            }

                            _cheminFichier = cheminAlternatif;
                            FallbackTimestampUtilise = true;
                        }
                    }

                    _workbook = partirDuTemplate
                        ? new XLWorkbook(templatePath)
                        : new XLWorkbook(_cheminFichier);
                }
                else
                {
                    _workbook = new XLWorkbook(templatePath);
                }

                var modFeuille = _workbook.Worksheet(NOM_MODELE);

                // --- 3. Création d'une nouvelle feuille (copie de ModFeuille) ---
                _nomFeuilleMesure = TrouverNomFeuilleUnique(config.TypeMesure);
                _feuilleMesure = modFeuille.CopyTo(_nomFeuilleMesure);

                // ModFeuille est cachée dans le template (convention métier — c'est juste un
                // modèle interne qui ne doit pas apparaître à l'utilisateur). Mais la copie
                // hérite de cet attribut Hidden — il faut donc forcer la visibilité de chaque
                // nouvelle feuille de mesure (1, 2, 3, … ou Freq1, Stab1, …).
                _feuilleMesure.Visibility = XLWorksheetVisibility.Visible;

                DeprotegerFeuille(_feuilleMesure);

                _feuilleMesure.Column("B").Width = 30;
                _feuilleMesure.Column("C").Width = 28;
                _feuilleMesure.Column("E").Width = 30;

                // --- 4. Clonage des zones nommées en sheet-scope sur la nouvelle feuille ---
                ClonerZonesNommeesPourNouvelleFeuille(modFeuille, _feuilleMesure);

                // --- 5. En-têtes adaptatifs (colonnes B/C/D + labels de lignes) ---
                var entetes = EnTetesMesureHelper.Pour(config.TypeMesure);
                _feuilleMesure.Cell("A7").SetValue(entetes.EnteteHeure);
                _feuilleMesure.Cell("B7").SetValue(entetes.EnteteMesuree);
                _feuilleMesure.Cell("C7").SetValue(entetes.EnteteReelle);
                _feuilleMesure.Cell("D7").SetValue(entetes.EnteteDelta);
                _feuilleMesure.Cell("B13").SetValue(entetes.LabelMoyenne);
                _feuilleMesure.Cell("B21").SetValue(entetes.LabelFreqRef);
                _feuilleMesure.Cell("B23").SetValue(entetes.LabelFreqCorr);
                _feuilleMesure.Cell("B25").SetValue(entetes.LabelIncertResol);
                _feuilleMesure.Cell("B31").SetValue(entetes.LabelIncertGlob);

                // --- 6. Métadonnées via zones nommées (sheet-scope) ---
                SetNamed("ZNNoFiche", numeroFI);
                SetNamed("ZNDate", DateTime.Now.ToString("dd/MM/yyyy"));
                SetNamed("ZNTypeMesure", EnTetesMesureHelper.LibelleType(config.TypeMesure));
                SetNamed("ZNFreqUtilise", NomAppareilDepuisCatalogue(config.IdModeleCatalogue));
                SetNamed("ZNRubidium",
                    rubidium.Designation + (rubidium.AvecGPS ? " (raccord GPS)" : " (raccord Allouis)"));
                SetNamed("ZNGate", EnTetesMesureHelper.LibelleGate(gateInscrite));
                SetNamed("ZNLibGate", EnTetesMesureHelper.LibelleGate(gateInscrite));
                SetNamed("ZNValGateSecondes", EnTetesMesureHelper.SecondesGate(gateInscrite));
                SetNamed("ZNModeMesure", config.ModeMesure == ModeMesure.Direct ? "Direct" : "Indirect");
                SetNamed("ZNCoeffMult", config.IndexMultiplicateur);
                SetNamed("ZNValFNominale", config.FNominale);
                SetNamed("ZNNbMesures", config.NbMesures);
                SetNamed("ZNIncertResol", config.Resolution);
                SetNamed("ZNIncertSup", config.IncertSupp);
                SetNamed("ZNFreqRef", rubidium.FrequenceMoyenne);

                // TODO : charger coefficients A/B et accréditation depuis SQL
                SetNamed("ZNCoeffA", 1e-10);
                SetNamed("ZNCoeffB", 5e-13);
                SetNamed("ZNNbMesAccredite", 30);
                SetNamed("ZNTempsMesureAccredite", 10);
            });
        }

        public async Task PreparerLignesMesureAsync(int nbMesures)
        {
            if (_feuilleMesure == null || _workbook == null || nbMesures <= 0) return;

            await Task.Run(() =>
            {
                // *** KEY FIX (identique à l'ancien AjouterResultatsAsync) ***
                // Le template a 2 lignes de mesures (9 et 10) et les labels commencent en 13.
                // Pour chaque mesure au-delà de la 2e, on insère une rangée AVANT ZNPointInsertion.
                if (nbMesures > 2)
                {
                    int pointInsertionRow = TrouverLigneZone("ZNPointInsertion") ?? 11;
                    _feuilleMesure.Row(pointInsertionRow).InsertRowsAbove(nbMesures - 2);
                }

                // Ajoute les formules Fréq. Réelle (col C) et delta (col D) pour les lignes
                // créées par InsertRowsAbove. Les cellules HEURE/mesure restent vides — elles
                // seront écrites en direct par ExcelInteropHost pendant la boucle de mesures.
                for (int i = 2; i < nbMesures; i++)
                {
                    int row = LIGNE_DEBUT_MESURES + i;
                    _feuilleMesure.Cell($"C{row}").FormulaA1 =
                        $"IF(ISBLANK(ZNCoeffMult),B{row},"
                        + $"(((B{row}-10000000)/(POWER(10,ZNCoeffMult)*10000000))+1)*ZNValFNominale)";
                    _feuilleMesure.Cell($"D{row}").FormulaA1 = $"C{row - 1}-C{row}";
                }
            });
        }

        public Task EcrireValeursBatchClosedXMLAsync(int ligneDebut, IList<(DateTime ts, double valeur)> mesures)
        {
            if (_feuilleMesure == null || mesures.Count == 0) return Task.CompletedTask;
            return Task.Run(() =>
            {
                for (int i = 0; i < mesures.Count; i++)
                {
                    int row = ligneDebut + i;
                    _feuilleMesure.Cell(row, 1).SetValue(mesures[i].ts.ToString("HH:mm:ss"));
                    _feuilleMesure.Cell(row, 2).SetValue(mesures[i].valeur);
                }
            });
        }

        public async Task EcrireStatsAsync(List<double> resultats)
        {
            if (_feuilleMesure == null || _workbook == null || resultats.Count == 0) return;

            await Task.Run(() =>
            {
                int n = resultats.Count;
                double moyenne = resultats.Average();
                double variance = 0;
                if (n >= 2)
                {
                    double m = moyenne;
                    double sum = 0;
                    foreach (var v in resultats) sum += (v - m) * (v - m);
                    variance = sum / (n - 1);
                }

                SetNamed("ZNFreqMoyReel", moyenne);
                SetNamed("ZNVariance", variance);
            });
        }

        public async Task<string> SauvegarderSurDisqueAsync()
        {
            if (_workbook == null) return string.Empty;
            await Task.Run(() =>
            {
                try { _workbook.SaveAs(_cheminFichier); }
                catch (IOException)
                {
                    throw new InvalidOperationException(
                        $"Impossible de sauvegarder « {Path.GetFileName(_cheminFichier)} » : "
                        + "le fichier est verrouillé. Fermez-le dans Excel et relancez la mesure.");
                }

                // Patche le lien vers Metrologo.xla dès la première sauvegarde pour que les formules
                // d'incertitude soient résolues dès l'ouverture du fichier par Excel Interop.
                try { PatcherLienMacroXLA(_cheminFichier, Preferences.CheminMacroXLA); }
                catch { /* best-effort */ }
            });
            return _cheminFichier;
        }

        public async Task RouvrirClasseurAsync()
        {
            if (string.IsNullOrEmpty(_cheminFichier) || string.IsNullOrEmpty(_nomFeuilleMesure))
                throw new InvalidOperationException(
                    "RouvrirClasseurAsync appelée avant InitialiserRapportAsync — pas de classeur à rouvrir.");

            await Task.Run(() =>
            {
                _workbook?.Dispose();
                _workbook = new XLWorkbook(_cheminFichier);

                // Récupère la feuille de mesure par son nom (créée à l'initialisation).
                _feuilleMesure = _workbook.Worksheets
                    .FirstOrDefault(w => string.Equals(w.Name, _nomFeuilleMesure, StringComparison.OrdinalIgnoreCase));

                if (_feuilleMesure == null)
                    throw new InvalidOperationException(
                        $"Feuille « {_nomFeuilleMesure} » introuvable après réouverture de « {Path.GetFileName(_cheminFichier)} ».");
            });
        }

        public async Task SauvegarderFinalAsync()
        {
            if (_workbook == null) return;
            await Task.Run(() =>
            {
                try { _workbook.SaveAs(_cheminFichier); }
                catch (IOException)
                {
                    throw new InvalidOperationException(
                        $"Impossible de sauvegarder « {Path.GetFileName(_cheminFichier)} » : "
                        + "le fichier est verrouillé. Fermez-le dans Excel et relancez la mesure.");
                }

                try { PatcherLienMacroXLA(_cheminFichier, Preferences.CheminMacroXLA); }
                catch { /* best-effort */ }
            });
        }

        /// <summary>
        /// Ajoute une ligne de récapitulatif Fréquence dans la feuille <c>Récap.</c>.
        /// Portage de <c>TfrmMain.MajRecapFreq</c> (F_Main.pas:2421) — la structure de Recap
        /// n'est pas modifiée, seules de nouvelles lignes de données sont insérées au point
        /// défini par la zone nommée <c>ZNRecapF_DebZone</c>.
        /// </summary>
        public async Task MettreAJourRecapFreqAsync(Mesure mesure)
        {
            if (_workbook == null || _feuilleMesure == null) return;
            string nomFeuille = _feuilleMesure.Name;

            await Task.Run(() => EcrireLigneRecap(
                nomFeuille,
                ZN_RECAPF_DEBZONE, LIGNE_FALLBACK_RECAPF,
                new[]
                {
                    $"='{nomFeuille}'!ZNFreqMoyReel",       // Col 1 : fréquence moyenne
                    $"='{nomFeuille}'!ZNLibGate",           // Col 2 : temps de mesure (libellé gate)
                    $"='{nomFeuille}'!ZNFreqCorr",          // Col 3 : fréquence corrigée
                    $"='{nomFeuille}'!ZNEcartType",         // Col 4 : écart-type
                    null,                                    // Col 5 : fréquence indiquée (valeur directe)
                    $"='{nomFeuille}'!ZNIncertResol",       // Col 6 : incertitude de résolution
                    $"='{nomFeuille}'!ZNIncertSup",         // Col 7 : incertitude supplémentaire
                    $"='{nomFeuille}'!ZNIncertAccreditee",  // Col 8 : incertitude accréditée
                    $"='{nomFeuille}'!ZNIncertGlobale"      // Col 9 : incertitude globale
                },
                colValeurDirecte: 5,
                valeurDirecte: mesure.SourceMesure == SourceMesure.Generateur
                    ? (object)"Géné."
                    : mesure.FNominale));
        }

        /// <summary>
        /// Ajoute une ligne dans la feuille <c>Récap.</c> du fichier Stabilité, en s'alignant
        /// sur la structure historique (cf. Stab1.xls) :
        /// <list type="bullet">
        ///   <item>8 colonnes : Temps (s), Fréq Moyenne, Écart type, Incertitude, Incert accréditée, Incert globale, Valeurs Maxi, Valeurs Mini</item>
        ///   <item>Insertion en ordre chronologique (gate 1 → L6, gate 2 → L7, …)</item>
        ///   <item>Cols 7-8 = formules locales <c>=C{n}+F{n}</c> et <c>=C{n}-F{n}</c> (moyenne ± incert. globale)</item>
        ///   <item>L1 mise à jour avec NumFI + date</item>
        /// </list>
        /// Le template Stab a déjà la ligne L5 comme template + L19 globaux (Max/Min).
        /// </summary>
        public async Task MettreAJourRecapStabAsync(Mesure mesure)
        {
            if (_workbook == null || _feuilleMesure == null) return;
            string nomFeuille = _feuilleMesure.Name;

            await Task.Run(() =>
            {
                if (!_workbook.Worksheets.Any(w => w.Name == NOM_RECAP)) return;
                var recap = _workbook.Worksheet(NOM_RECAP);
                DeprotegerFeuille(recap);

                // En-tête fiche : NumFI + date — idempotent à chaque itération de gate.
                recap.Cell("B1").SetValue(mesure.NumFI);
                recap.Cell("D1").SetValue(DateTime.Now.ToString("dd/MM/yyyy"));

                // Numéro de la gate dans la séquence (1, 2, 3…). Le nommage des feuilles de
                // mesure étant numérique pur en Stab, on parse directement le nom — sinon
                // fallback : on compte les lignes déjà remplies dans la zone de données.
                int numero;
                if (!int.TryParse(nomFeuille, out numero))
                {
                    numero = CompterLignesStabRemplies(recap) + 1;
                }

                int ligne = 5 + numero;  // L6 = 1ère gate balayée, L7 = 2ème, etc.

                recap.Cell(ligne, 1).FormulaA1 = $"='{nomFeuille}'!ZNValGateSecondes";
                recap.Cell(ligne, 2).FormulaA1 = $"='{nomFeuille}'!ZNFreqMoyReel";
                recap.Cell(ligne, 3).FormulaA1 = $"='{nomFeuille}'!ZNEcartType";
                recap.Cell(ligne, 4).FormulaA1 = $"='{nomFeuille}'!ZNIncertEcartType";
                recap.Cell(ligne, 5).FormulaA1 = $"='{nomFeuille}'!ZNIncertAccreditee";
                recap.Cell(ligne, 6).FormulaA1 = $"='{nomFeuille}'!ZNIncertGlobale";
                recap.Cell(ligne, 7).FormulaA1 = $"=C{ligne}+F{ligne}";
                recap.Cell(ligne, 8).FormulaA1 =
                    $"=IF(ISBLANK(C{ligne}),,IF((C{ligne}-F{ligne})<=0,0,C{ligne}-F{ligne}))";
            });
        }

        /// <summary>
        /// Compte le nombre de lignes de données déjà remplies dans la zone Stab (L6+).
        /// Une ligne est considérée remplie si sa colonne A (Temps) contient quelque chose.
        /// </summary>
        private static int CompterLignesStabRemplies(IXLWorksheet recap)
        {
            int compteur = 0;
            for (int row = 6; row <= 18; row++)
            {
                var cell = recap.Cell(row, 1);
                if (!cell.IsEmpty()) compteur++;
            }
            return compteur;
        }

        // Ligne d'entête des colonnes dans le template Récap. (observé dans METROLOGO.xltm).
        // Toute nouvelle mesure est insérée juste en dessous (row 6) pour avoir "newest on top".
        private const int LIGNE_ENTETE_RECAP = 5;

        /// <summary>
        /// Insère une nouvelle ligne en tête de la zone de données Récap. et la remplit avec
        /// des formules cross-sheet vers la feuille de mesure. Les anciennes mesures descendent
        /// d'une unité ; la plus récente reste ainsi toujours au sommet.
        /// </summary>
        private void EcrireLigneRecap(
            string nomFeuilleMesure,
            string zoneDebZone,
            int ligneFallback,
            string?[] formulesParColonne,
            int colValeurDirecte = -1,
            object? valeurDirecte = null)
        {
            if (_workbook == null) return;
            if (!_workbook.Worksheets.Any(w => w.Name == NOM_RECAP)) return;

            var recap = _workbook.Worksheet(NOM_RECAP);

            DeprotegerFeuille(recap);

            // 1. Nettoyage des lignes "fantômes" héritées du template : ces lignes contiennent
            //    des formules =[0]!ZNxxx qui pointent vers ModFeuille (toujours vide) et
            //    produisent une plage de zéros sans intérêt juste sous l'entête.
            NettoyerLignesGhost(recap);

            // 2. Insertion de la nouvelle ligne directement sous l'entête — la plus récente en haut.
            int nouvelleLigne = LIGNE_ENTETE_RECAP + 1;
            recap.Row(nouvelleLigne).InsertRowsAbove(1);

            // 3. Remplissage des colonnes
            for (int i = 0; i < formulesParColonne.Length; i++)
            {
                int col = i + 1;
                if (col == colValeurDirecte && valeurDirecte != null)
                {
                    recap.Cell(nouvelleLigne, col).SetValue(XLCellValue.FromObject(valeurDirecte));
                    continue;
                }
                var formule = formulesParColonne[i];
                if (!string.IsNullOrEmpty(formule))
                    recap.Cell(nouvelleLigne, col).FormulaA1 = formule;
            }
        }

        /// <summary>
        /// Supprime les lignes de la zone de données qui ne contiennent que des formules
        /// <c>=[0]!ZNxxx</c> (placeholders du template pointant vers ModFeuille, toujours vide).
        /// Le <c>[0]!</c> est un indicateur sans ambiguïté de ces lignes issues du template xltm.
        /// Itération descendante pour ne pas décaler les indices pendant la suppression.
        /// </summary>
        private static void NettoyerLignesGhost(IXLWorksheet recap)
        {
            int derniereLigne = recap.LastRowUsed()?.RowNumber() ?? LIGNE_ENTETE_RECAP;
            for (int row = derniereLigne; row > LIGNE_ENTETE_RECAP; row--)
            {
                if (LigneEstGhost(recap, row))
                    recap.Row(row).Delete();
            }
        }

        private static bool LigneEstGhost(IXLWorksheet recap, int row)
        {
            foreach (var cell in recap.Row(row).CellsUsed())
            {
                if (!cell.HasFormula) continue;
                var f = cell.FormulaA1 ?? string.Empty;
                if (f.Contains("[0]!")) return true;
            }
            return false;
        }


        /// <summary>
        /// Remplace les caractères interdits par Windows dans les noms de fichier/dossier
        /// (<c>&lt; &gt; : " / \ | ? *</c> + caractères de contrôle) par un underscore. Utilisé
        /// pour transformer un numéro FI saisi par l'utilisateur en nom de dossier sûr.
        /// </summary>
        private static string SanitizerNomFichier(string nom)
        {
            if (string.IsNullOrWhiteSpace(nom)) return "sans-nom";

            var invalides = new HashSet<char>(Path.GetInvalidFileNameChars());
            var sb = new System.Text.StringBuilder(nom.Length);
            foreach (var c in nom)
            {
                sb.Append(invalides.Contains(c) ? '_' : c);
            }
            string resultat = sb.ToString().Trim(' ', '.');
            return string.IsNullOrEmpty(resultat) ? "sans-nom" : resultat;
        }

        /// <summary>
        /// Réécrit le Target de la relation externe dans le fichier .xlsm pour qu'il
        /// pointe vers le chemin configuré du fichier Metrologo.xla.
        /// </summary>
        private static void PatcherLienMacroXLA(string xlsmPath, string xlaPath)
        {
            const string relPath = "xl/externalLinks/_rels/externalLink1.xml.rels";

            // URI file:/// avec backslashes → slashes pour compatibilité XML
            string uri = "file:///" + xlaPath.Replace('\\', '/');

            using var zip = ZipFile.Open(xlsmPath, ZipArchiveMode.Update);
            var entry = zip.GetEntry(relPath);
            if (entry == null) return;

            string contenu;
            using (var reader = new StreamReader(entry.Open()))
            {
                contenu = reader.ReadToEnd();
            }

            // Remplace chaque Target="...Metrologo.xla" par le chemin configuré
            var rewritten = Regex.Replace(contenu,
                @"Target=""[^""]*Metrologo\.xla""",
                $@"Target=""{uri}""",
                RegexOptions.IgnoreCase);

            if (rewritten == contenu) return; // rien à changer

            entry.Delete();
            var newEntry = zip.CreateEntry(relPath);
            using var writer = new StreamWriter(newEntry.Open());
            writer.Write(rewritten);
        }

        public void FermerExcel()
        {
            _workbook?.Dispose();
            _workbook = null;
            _feuilleMesure = null;
        }

        // ---------- Utilitaires ----------

        // Mots de passe connus appliqués à ModFeuille dans les templates métier.
        private static readonly string[] _motsDePasseFeuille = { "METROL", "metrol" };

        private static void DeprotegerFeuille(IXLWorksheet feuille)
        {
            try { feuille.Unprotect(); return; }
            catch { /* feuille protégée par mot de passe — on essaie ci-dessous */ }

            foreach (var mdp in _motsDePasseFeuille)
            {
                try { feuille.Unprotect(mdp); return; }
                catch { /* mauvais mot de passe — on essaie le suivant */ }
            }

            throw new InvalidOperationException(
                $"Impossible de déprotéger la feuille « {feuille.Name} » : "
                + "aucun mot de passe connu ne correspond. Vérifiez le template.");
        }

        private void SetNamed(string name, object value)
        {
            if (_feuilleMesure == null) return;
            try
            {
                if (_feuilleMesure.DefinedNames.TryGetValue(name, out var defName))
                {
                    defName.Ranges.First().FirstCell().SetValue(XLCellValue.FromObject(value));
                }
            }
            catch { /* zone nommée absente → silencieux */ }
        }

        private static string NomAppareilDepuisCatalogue(string idModele)
        {
            if (string.IsNullOrEmpty(idModele)) return string.Empty;
            var modele = CatalogueAppareilsService.Instance.Modeles
                .FirstOrDefault(m => m.Id == idModele);
            return modele?.Nom ?? idModele;
        }

        private int? TrouverLigneZone(string name)
        {
            if (_feuilleMesure == null) return null;
            try
            {
                if (_feuilleMesure.DefinedNames.TryGetValue(name, out var defName))
                {
                    return defName.Ranges.First().FirstCell().Address.RowNumber;
                }
            }
            catch { }
            return null;
        }

        private static bool FichierEstVerrouille(string chemin)
        {
            try
            {
                using var fs = File.Open(chemin, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                return false;
            }
            catch (IOException)
            {
                return true;
            }
        }

        private string TrouverNomFeuilleUnique(TypeMesure type)
        {
            if (type == TypeMesure.FreqAvantInterv) return "F_Avant_Interv";
            if (type == TypeMesure.FreqFinale) return "F_Finale";

            // Stabilité : nommage numérique pur (1, 2, 3…) — aligné sur la convention
            // historique du Delphi/Stab1.xls. La feuille Récap. attend ce format pour
            // ses formules cross-sheet ='1'!ZN…, ='2'!ZN…, etc.
            if (type == TypeMesure.Stabilite)
            {
                int n = 1;
                var existantsStab = _workbook!.Worksheets.Select(w => w.Name)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
                while (existantsStab.Contains(n.ToString())) n++;
                return n.ToString();
            }

            string prefixe = type switch
            {
                TypeMesure.Frequence => "Freq",
                TypeMesure.Interval => "Interv",
                TypeMesure.TachyContact => "TachyC",
                TypeMesure.Stroboscope => "Strobo",
                _ => "Mesure"
            };

            int idx = 1;
            var existants = _workbook!.Worksheets.Select(w => w.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            while (existants.Contains($"{prefixe}{idx}")) idx++;
            return $"{prefixe}{idx}";
        }

        private void ClonerZonesNommeesPourNouvelleFeuille(IXLWorksheet source, IXLWorksheet dest)
        {
            if (_workbook == null) return;

            foreach (var defName in _workbook.DefinedNames.ToList())
            {
                try
                {
                    if (!defName.RefersTo.Contains(source.Name + "!", StringComparison.OrdinalIgnoreCase))
                        continue;

                    string newRefersTo = defName.RefersTo
                        .Replace(source.Name + "!", dest.Name + "!",
                                 StringComparison.OrdinalIgnoreCase);

                    if (!dest.DefinedNames.Contains(defName.Name))
                    {
                        dest.DefinedNames.Add(defName.Name, newRefersTo);
                    }
                }
                catch { /* ignore #REF! */ }
            }
        }
    }
}
