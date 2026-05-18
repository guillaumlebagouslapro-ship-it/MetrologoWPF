using System.Collections.Generic;
using System.IO;
using System.Windows;
using Metrologo.Services;
using Metrologo.Services.Journal;
using Wpf.Ui.Controls;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Views
{
    /// <summary>
    /// Fenêtre affichée au premier démarrage de l'app sur un poste pour demander à
    /// l'utilisateur le chemin local de sauvegarde des rapports Excel. L'utilisateur
    /// peut reporter — dans ce cas un bandeau d'alerte reste visible dans MainWindow
    /// jusqu'à la configuration.
    /// </summary>
    public partial class PremierDemarrageCheminLocalWindow : FluentWindow
    {
        public PremierDemarrageCheminLocalWindow()
        {
            InitializeComponent();
        }

        private void OnParcourir(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFolderDialog
            {
                Title = "Choisir un dossier local de sauvegarde",
                InitialDirectory = Directory.Exists(TbChemin.Text) ? TbChemin.Text : "C:\\"
            };
            if (dlg.ShowDialog() == true)
            {
                TbChemin.Text = dlg.FolderName;
            }
        }

        private void OnValider(object sender, RoutedEventArgs e)
        {
            string chemin = TbChemin.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(chemin))
            {
                System.Windows.MessageBox.Show("Saisis ou choisis un dossier, ou clique « Plus tard ».",
                    "Chemin requis", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            // Vérifie qu'on peut écrire dedans (création + test write/delete).
            try
            {
                Directory.CreateDirectory(chemin);
                string testFile = Path.Combine(chemin, $".test_{System.Guid.NewGuid():N}.tmp");
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"Impossible d'écrire dans « {chemin} » :\n{ex.Message}\n\nChoisis un autre dossier.",
                    "Accès refusé", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return;
            }

            // Persiste le chemin dans paths.config.json (préserve les autres overrides).
            CheminsMetrologo.ChargerConfigChemins();
            var overrides = new Dictionary<string, string>();
            foreach (var cle in new[] {
                nameof(CheminsMetrologo.Incertitudes),
                nameof(CheminsMetrologo.Presets),
                nameof(CheminsMetrologo.Catalogues),
                nameof(CheminsMetrologo.ArchivesLogs) })
            {
                if (CheminsMetrologo.EstSurcharge(cle))
                {
                    // On relit via la propriété typée pour récupérer la valeur courante.
                    var prop = typeof(CheminsMetrologo).GetProperty(cle);
                    if (prop != null)
                        overrides[cle] = (string)prop.GetValue(null)!;
                }
            }
            overrides[nameof(CheminsMetrologo.MesuresLocal)] = chemin;
            CheminsMetrologo.EnregistrerConfigChemins(overrides);

            JournalLog.Info(CategorieLog.Administration, "CHEMIN_LOCAL_INITIALISE",
                $"Chemin local de sauvegarde configuré au premier démarrage : {chemin}");

            DialogResult = true;
            Close();
        }

        private void OnReporter(object sender, RoutedEventArgs e)
        {
            JournalLog.Warn(CategorieLog.Administration, "CHEMIN_LOCAL_REPORTE",
                "Configuration du chemin local reportée — bandeau d'alerte actif jusqu'à configuration.");
            DialogResult = false;
            Close();
        }
    }
}
