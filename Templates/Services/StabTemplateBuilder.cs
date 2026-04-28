using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;

namespace Metrologo.Services
{
    /// <summary>
    /// Construit <c>METROLOGO_Stab.xltm</c> à partir du template Fréquence existant.
    /// On garde la même <c>ModFeuille</c> (mêmes zones nommées de mesure : ZNFreqMoyReel,
    /// ZNVariance, ZNEcartType, ZNIncertResol, ZNIncertSup, ZNIncertAccreditee,
    /// ZNIncertGlobale, etc.) et on remplace uniquement la <c>Récap.</c> par la structure
    /// historique Stabilité (8 colonnes, 1 ligne par gate balayée, formules Max/Mini
    /// = moyenne ± incertitude globale).
    ///
    /// Appelé au démarrage de l'application : si le fichier cible existe déjà, ne fait rien.
    /// </summary>
    public static class StabTemplateBuilder
    {
        private const string NomRecap = "Récap.";

        public static void EnsureExists(string templateFreqPath, string templateStabPath)
        {
            if (File.Exists(templateStabPath)) return;
            if (!File.Exists(templateFreqPath))
                throw new FileNotFoundException(
                    $"Template Fréquence introuvable : {templateFreqPath}", templateFreqPath);

            File.Copy(templateFreqPath, templateStabPath, overwrite: true);

            using var wb = new XLWorkbook(templateStabPath);
            var recap = wb.Worksheet(NomRecap);

            // Déprotège (mots de passe historiques connus).
            try { recap.Unprotect(); }
            catch { try { recap.Unprotect("METROL"); } catch { try { recap.Unprotect("metrol"); } catch { } } }

            // 1. Supprime les zones nommées workbook-scope spécifiques à Fréquence/Dérive.
            //    On garde les zones partagées (ZNDate, ZNNoFiche, ZNFreqMoyReel, etc.) qui
            //    pointent sur ModFeuille — elles sont valides pour les deux types de mesure.
            var aSupprimer = wb.DefinedNames
                .Where(n =>
                    n.Name.StartsWith("ZNNoFicheRecap", StringComparison.OrdinalIgnoreCase) ||
                    n.Name.StartsWith("ZNDateRecap", StringComparison.OrdinalIgnoreCase) ||
                    n.Name.StartsWith("ZNRecapF_", StringComparison.OrdinalIgnoreCase) ||
                    n.Name.StartsWith("ZNRecapD_", StringComparison.OrdinalIgnoreCase) ||
                    n.Name.StartsWith("ZNRecapS_", StringComparison.OrdinalIgnoreCase) ||
                    n.Name.StartsWith("ZNGrandeValeur", StringComparison.OrdinalIgnoreCase) ||
                    n.Name.StartsWith("ZNPetiteValeur", StringComparison.OrdinalIgnoreCase))
                .Select(n => n.Name)
                .ToList();
            foreach (var nom in aSupprimer)
            {
                try { wb.DefinedNames.Delete(nom); } catch { /* ignore */ }
            }

            // 2. Vide la feuille Récap. (cellules + formats).
            recap.Clear();

            // 3. Écrit la structure Stab historique
            // Ligne 1 : en-tête fiche
            recap.Cell("A1").SetValue("Fiche n° :");
            recap.Cell("A1").Style.Font.Bold = true;
            recap.Cell("C1").SetValue("Le");
            recap.Cell("C1").Style.Font.Bold = true;

            // Ligne 3 : en-têtes des colonnes Stab
            string[] entetes =
            {
                "Temps (s)", "Fréq. Moyenne (Hz)", "Ecart type", "Incertitude",
                "Incert. accréditée", "Incert. globale (Hz)", "Valeurs Maxi.", "Valeurs Mini."
            };
            for (int i = 0; i < entetes.Length; i++)
            {
                var cell = recap.Cell(3, i + 1);
                cell.SetValue(entetes[i]);
                cell.Style.Font.Bold = true;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            }

            // Ligne 5 : ligne template (sera dupliquée à chaque nouvelle gate insérée).
            // Les formules Maxi/Mini se réfèrent à C5/F5 et restent cohérentes après duplication
            // (ClosedXML translate les références relatives quand on insère).
            recap.Cell("G5").FormulaA1 = "=C5+F5";
            recap.Cell("H5").FormulaA1 = "=IF(ISBLANK(C5),,IF((C5-F5)<=0,0,C5-F5))";

            // Ligne 19 : Max / Min global sur la zone des données (suffisant pour 13 gates).
            recap.Cell("A19").SetValue("Globale");
            recap.Cell("A19").Style.Font.Bold = true;
            recap.Cell("G19").FormulaA1 = "=MAX(G6:G15)";
            recap.Cell("H19").FormulaA1 = "=MIN(H6:H15)";

            // 4. Zones nommées workbook-scope spécifiques à Stab
            wb.DefinedNames.Add("ZNNoFicheRecapStab", $"{NomRecap}!$B$1");
            wb.DefinedNames.Add("ZNDateRecapStab",    $"{NomRecap}!$D$1");
            wb.DefinedNames.Add("ZNRecapS_DebZone",   $"{NomRecap}!$A$6");
            wb.DefinedNames.Add("ZNRecapS_Ligne0",    $"{NomRecap}!$5:$5");
            wb.DefinedNames.Add("ZNGrandeValeur",     $"{NomRecap}!$G$19");
            wb.DefinedNames.Add("ZNPetiteValeur",     $"{NomRecap}!$H$19");

            // 5. Largeur des colonnes (lisibilité)
            for (int c = 1; c <= 8; c++) recap.Column(c).Width = 22;

            wb.Save();
        }
    }
}
