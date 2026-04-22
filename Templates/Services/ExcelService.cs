using ClosedXML.Excel;
using Metrologo.Models;
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
        Task InitialiserRapportAsync(string numeroFI, Mesure configuration, Rubidium rubidium);
        Task AjouterResultatsAsync(List<double> resultats);
        Task MettreAJourRecapFreqAsync(Mesure mesure);
        Task MettreAJourRecapStabAsync(Mesure mesure);
        Task SauvegarderEtOuvrirAsync();
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

        public async Task InitialiserRapportAsync(string numeroFI, Mesure config, Rubidium rubidium)
        {
            await Task.Run(() =>
            {
                // --- 1. Détermination du dossier et du fichier (un fichier par FI) ---
                string dossier = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "Metrologo", numeroFI);
                Directory.CreateDirectory(dossier);
                _cheminFichier = Path.Combine(dossier, $"Mesures_{numeroFI}.xlsm");

                // --- 2. Ouverture : fichier existant ou copie du template ---
                string templatePath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "Templates", "METROLOGO.xltm");

                if (File.Exists(_cheminFichier))
                {
                    if (FichierEstVerrouille(_cheminFichier))
                    {
                        WindowHelper.FermerFenetresExcel(Path.GetFileName(_cheminFichier));
                        var sw = System.Diagnostics.Stopwatch.StartNew();
                        while (FichierEstVerrouille(_cheminFichier) && sw.ElapsedMilliseconds < 5000)
                        {
                            System.Threading.Thread.Sleep(150);
                        }
                        if (FichierEstVerrouille(_cheminFichier))
                        {
                            throw new InvalidOperationException(
                                $"Le fichier « {Path.GetFileName(_cheminFichier)} » est toujours ouvert dans Excel."
                                + Environment.NewLine + Environment.NewLine
                                + "Fermez Excel manuellement (y compris tout dialogue « Enregistrer ? ») puis réessayez.");
                        }
                    }
                    _workbook = new XLWorkbook(_cheminFichier);
                }
                else
                {
                    _workbook = new XLWorkbook(templatePath);
                }

                var modFeuille = _workbook.Worksheet(NOM_MODELE);

                // --- 3. Création d'une nouvelle feuille (copie de ModFeuille) ---
                string nomFeuille = TrouverNomFeuilleUnique(config.TypeMesure);
                _feuilleMesure = modFeuille.CopyTo(nomFeuille);

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
                SetNamed("ZNFreqUtilise", EnTetesMesureHelper.NomAppareil(config.Frequencemetre));
                SetNamed("ZNRubidium",
                    rubidium.Designation + (rubidium.AvecGPS ? " (raccord GPS)" : " (raccord Allouis)"));
                SetNamed("ZNGate", EnTetesMesureHelper.LibelleGate(config.GateIndex));
                SetNamed("ZNLibGate", EnTetesMesureHelper.LibelleGate(config.GateIndex));
                SetNamed("ZNValGateSecondes", EnTetesMesureHelper.SecondesGate(config.GateIndex));
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

        public async Task AjouterResultatsAsync(List<double> resultats)
        {
            if (_feuilleMesure == null || _workbook == null || resultats.Count == 0) return;

            await Task.Run(() =>
            {
                int n = resultats.Count;

                // *** KEY FIX ***
                // Le template a 2 lignes de mesures (9 et 10) et les labels commencent en 13.
                // Pour chaque mesure au-delà de la 2e, on insère une rangée AVANT
                // ZNPointInsertion (=A11) pour pousser les labels vers le bas, comme le
                // faisait l'ancien Delphi.
                if (n > 2)
                {
                    int pointInsertionRow = TrouverLigneZone("ZNPointInsertion") ?? 11;
                    _feuilleMesure.Row(pointInsertionRow).InsertRowsAbove(n - 2);
                }

                // Écrit les valeurs : HEURE en A, mesure brute en B.
                var maintenant = DateTime.Now;
                for (int i = 0; i < n; i++)
                {
                    int row = LIGNE_DEBUT_MESURES + i;
                    _feuilleMesure.Cell($"A{row}").SetValue(
                        maintenant.AddSeconds(i).ToString("HH:mm:ss"));
                    _feuilleMesure.Cell($"B{row}").SetValue(resultats[i]);

                    // Pour les lignes créées par InsertRowsAbove, pas de formule héritée
                    // → on ajoute la formule F réelle (col C) et delta (col D) à la main.
                    if (i >= 2)
                    {
                        _feuilleMesure.Cell($"C{row}").FormulaA1 =
                            $"IF(ISBLANK(ZNCoeffMult),B{row},"
                            + $"(((B{row}-10000000)/(POWER(10,ZNCoeffMult)*10000000))+1)*ZNValFNominale)";
                        _feuilleMesure.Cell($"D{row}").FormulaA1 = $"C{row - 1}-C{row}";
                    }
                }

                // Moyenne et variance — écrites via zones nommées qui ont suivi le shift.
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
        /// Ajoute une ligne de récapitulatif Stabilité dans la feuille <c>Récap.</c>.
        /// Portage de <c>TfrmMain.MajRecapStab</c> (F_Main.pas:2470).
        /// </summary>
        public async Task MettreAJourRecapStabAsync(Mesure mesure)
        {
            if (_workbook == null || _feuilleMesure == null) return;
            string nomFeuille = _feuilleMesure.Name;

            await Task.Run(() => EcrireLigneRecap(
                nomFeuille,
                ZN_RECAPS_DEBZONE, LIGNE_FALLBACK_RECAPS,
                new[]
                {
                    $"='{nomFeuille}'!ZNValGateSecondes",   // Col 1 : gate en secondes
                    $"='{nomFeuille}'!ZNFreqMoyReel",       // Col 2 : fréquence moyenne
                    $"='{nomFeuille}'!ZNEcartType",         // Col 3 : écart-type
                    $"='{nomFeuille}'!ZNIncertEcartType",   // Col 4 : incertitude sur l'écart-type
                    $"='{nomFeuille}'!ZNIncertAccreditee",  // Col 5 : incertitude accréditée
                    $"='{nomFeuille}'!ZNIncertGlobale"      // Col 6 : incertitude globale
                }));
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

        public async Task SauvegarderEtOuvrirAsync()
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

                // Patche la référence externe vers Metrologo.xla selon le chemin configuré.
                // Nécessaire pour que les macros Cal_ecart_type, Cal_freq_corrigee, etc.
                // soient résolues sur le poste utilisateur.
                try { PatcherLienMacroXLA(_cheminFichier, Preferences.CheminMacroXLA); }
                catch { /* best-effort — Excel affichera #NOM? si échec */ }

                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = _cheminFichier,
                    UseShellExecute = true
                });
            });
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

            string prefixe = type switch
            {
                TypeMesure.Frequence => "Freq",
                TypeMesure.Stabilite => "Stab",
                TypeMesure.Interval => "Interv",
                TypeMesure.TachyContact => "TachyC",
                TypeMesure.Stroboscope => "Strobo",
                _ => "Mesure"
            };

            int n = 1;
            var existants = _workbook!.Worksheets.Select(w => w.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            while (existants.Contains($"{prefixe}{n}")) n++;
            return $"{prefixe}{n}";
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
