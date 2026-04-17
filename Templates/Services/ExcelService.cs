using ClosedXML.Excel;
using Metrologo.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Metrologo.Services
{
    public interface IExcelService
    {
        Task InitialiserRapportAsync(string numeroFI, Mesure configuration);
        Task AjouterResultatsAsync(List<double> resultats);
        Task SauvegarderEtOuvrirAsync();
        void FermerExcel();
    }

    public class ExcelService : IExcelService
    {
        private XLWorkbook? _workbook;
        private IXLWorksheet? _feuilleMesure;
        private string _cheminFichierTravail = string.Empty;

        public async Task InitialiserRapportAsync(string numeroFI, Mesure config)
        {
            await Task.Run(() =>
            {
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "METROLOGO.xltm");

                // LE FIX EST ICI : On sauvegarde en .xlsm (avec macros) et non en .xlsx
                _cheminFichierTravail = Path.Combine(Path.GetTempPath(), $"Rapport_{numeroFI}_{DateTime.Now:yyyyMMddHHmmss}.xlsm");

                _workbook = new XLWorkbook(templatePath);
                _feuilleMesure = _workbook.Worksheet(1);

                SetNamedRangeValue("ZNNoFiche", numeroFI);
                SetNamedRangeValue("ZNDate", DateTime.Now.ToString("dd/MM/yyyy"));
            });
        }

        public async Task AjouterResultatsAsync(List<double> resultats)
        {
            if (_feuilleMesure == null || _workbook == null) return;

            await Task.Run(() =>
            {
                IXLCell? startCell = null;
                string nomZone = "ZNMESURE1";

                if (_workbook.DefinedNames.TryGetValue(nomZone, out var definedName))
                {
                    startCell = definedName.Ranges.First().FirstCell();
                }
                else if (_feuilleMesure.DefinedNames.TryGetValue(nomZone, out var sheetName))
                {
                    startCell = sheetName.Ranges.First().FirstCell();
                }

                if (startCell == null)
                {
                    throw new Exception($"La zone nommée '{nomZone}' est introuvable.");
                }

                int rowOffset = 0;
                foreach (var val in resultats)
                {
                    var currentCell = startCell.CellBelow(rowOffset);
                    currentCell.SetValue(DateTime.Now.ToString("HH:mm:ss"));
                    currentCell.CellRight().SetValue(val);
                    rowOffset++;
                }

                // Nous laissons volontairement le vrai logiciel Excel faire les calculs
                // à l'ouverture du fichier pour éviter les plantages.
            });
        }

        public async Task SauvegarderEtOuvrirAsync()
        {
            if (_workbook != null)
            {
                await Task.Run(() =>
                {
                    _workbook.SaveAs(_cheminFichierTravail);
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = _cheminFichierTravail,
                        UseShellExecute = true
                    });
                });
            }
        }

        public void FermerExcel() => _workbook?.Dispose();

        private void SetNamedRangeValue(string rangeName, object value)
        {
            if (_workbook != null && _workbook.DefinedNames.Contains(rangeName))
            {
                _workbook.Range(rangeName).FirstCell().SetValue(XLCellValue.FromObject(value));
            }
        }
    }
}