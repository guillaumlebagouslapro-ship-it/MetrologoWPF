using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Services;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// VM de la fenêtre admin "Chemins de stockage". Permet de surcharger les chemins
    /// par défaut (locaux, dans <c>%LocalAppData%\Metrologo\</c>) par des emplacements
    /// réseau partagés entre tous les postes du site. Persistance dans
    /// <c>Configuration\paths.config.json</c> (cf. <see cref="CheminsMetrologo"/>).
    /// </summary>
    public partial class CheminsStockageViewModel : ObservableObject
    {
        // ---------- Chemins éditables ----------

        [ObservableProperty] private string _cheminIncertitudes = string.Empty;
        [ObservableProperty] private string _cheminPresets = string.Empty;
        [ObservableProperty] private string _cheminCatalogues = string.Empty;
        [ObservableProperty] private string _cheminArchivesLogs = string.Empty;
        [ObservableProperty] private string _cheminMesuresLocal = string.Empty;

        [ObservableProperty] private string _statut = "Prêt — laisse un champ vide pour utiliser le chemin local par défaut.";

        // ---------- Infos non-éditables (toujours locales) ----------

        public string CheminConfiguration => CheminsMetrologo.Configuration;
        public string CheminCache => CheminsMetrologo.Cache;
        public string CheminFichierConfig => CheminsMetrologo.FichierPathsConfig;

        public CheminsStockageViewModel()
        {
            // Recharge depuis paths.config.json pour récupérer les overrides actuels.
            CheminsMetrologo.ChargerConfigChemins();

            // Pré-remplit les TextBox avec les valeurs actuelles SI surchargées,
            // sinon laisse vide (= utilise le défaut local).
            _cheminIncertitudes = CheminsMetrologo.EstSurcharge(nameof(CheminsMetrologo.Incertitudes))
                ? CheminsMetrologo.Incertitudes : string.Empty;
            _cheminPresets = CheminsMetrologo.EstSurcharge(nameof(CheminsMetrologo.Presets))
                ? CheminsMetrologo.Presets : string.Empty;
            _cheminCatalogues = CheminsMetrologo.EstSurcharge(nameof(CheminsMetrologo.Catalogues))
                ? CheminsMetrologo.Catalogues : string.Empty;
            _cheminArchivesLogs = CheminsMetrologo.EstSurcharge(nameof(CheminsMetrologo.ArchivesLogs))
                ? CheminsMetrologo.ArchivesLogs : string.Empty;
            _cheminMesuresLocal = CheminsMetrologo.MesuresLocal;
        }

        public Action<bool>? CloseAction { get; set; }

        // ---------- Commandes ----------

        [RelayCommand]
        private void ParcourirIncertitudes() => CheminIncertitudes = ParcourirDossier(CheminIncertitudes) ?? CheminIncertitudes;

        [RelayCommand]
        private void ParcourirPresets() => CheminPresets = ParcourirDossier(CheminPresets) ?? CheminPresets;

        [RelayCommand]
        private void ParcourirCatalogues() => CheminCatalogues = ParcourirDossier(CheminCatalogues) ?? CheminCatalogues;

        [RelayCommand]
        private void ParcourirArchivesLogs() => CheminArchivesLogs = ParcourirDossier(CheminArchivesLogs) ?? CheminArchivesLogs;

        [RelayCommand]
        private void ParcourirMesuresLocal() => CheminMesuresLocal = ParcourirDossier(CheminMesuresLocal) ?? CheminMesuresLocal;

        [RelayCommand]
        private void Tester()
        {
            var resultats = new List<string>();
            TesterDossier("Incertitudes", CheminIncertitudes, resultats);
            TesterDossier("Presets", CheminPresets, resultats);
            TesterDossier("Catalogues", CheminCatalogues, resultats);
            TesterDossier("Archives logs", CheminArchivesLogs, resultats);
            TesterDossier("Mesures (copie locale)", CheminMesuresLocal, resultats);

            string msg = string.Join(Environment.NewLine, resultats);
            MessageBox.Show(msg, "Test des chemins", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        [RelayCommand]
        private void Enregistrer()
        {
            try
            {
                var overrides = new Dictionary<string, string>
                {
                    [nameof(CheminsMetrologo.Incertitudes)] = CheminIncertitudes ?? string.Empty,
                    [nameof(CheminsMetrologo.Presets)]      = CheminPresets ?? string.Empty,
                    [nameof(CheminsMetrologo.Catalogues)]   = CheminCatalogues ?? string.Empty,
                    [nameof(CheminsMetrologo.ArchivesLogs)] = CheminArchivesLogs ?? string.Empty,
                    [nameof(CheminsMetrologo.MesuresLocal)] = CheminMesuresLocal ?? string.Empty
                };
                CheminsMetrologo.EnregistrerConfigChemins(overrides);

                JournalLog.Info(CategorieLog.Administration, "CHEMINS_SAUVE",
                    "Configuration des chemins de stockage mise à jour.");

                MessageBox.Show(
                    "Configuration enregistrée.\n\n"
                  + "⚠ Redémarre l'application pour que les nouveaux chemins soient pris en "
                  + "compte par TOUS les services (certains caches sont initialisés au démarrage).",
                    "Sauvegarde OK", MessageBoxButton.OK, MessageBoxImage.Information);

                CloseAction?.Invoke(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Sauvegarde échouée : {ex.Message}",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
                JournalLog.Erreur(CategorieLog.Administration, "CHEMINS_SAUVE_ERR",
                    $"Échec sauvegarde paths.config.json : {ex.Message}");
            }
        }

        [RelayCommand]
        private void Reinitialiser()
        {
            var conf = MessageBox.Show(
                "Vider tous les chemins surchargés ?\n\n"
              + "L'application utilisera alors les chemins par défaut locaux "
              + "(%LocalAppData%\\Metrologo\\…).",
                "Réinitialiser", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (conf != MessageBoxResult.Yes) return;

            CheminIncertitudes = string.Empty;
            CheminPresets = string.Empty;
            CheminCatalogues = string.Empty;
            CheminArchivesLogs = string.Empty;
            CheminMesuresLocal = string.Empty;
            Statut = "Champs vidés. Clique « Enregistrer » pour appliquer.";
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);

        // ---------- Helpers ----------

        private static string? ParcourirDossier(string cheminInitial)
        {
            // Ouvre un browser de dossier. Utilise le dialog standard via OpenFileDialog en
            // sélection multi-fichier pour atteindre le dossier (workaround WPF natif —
            // pas de FolderBrowserDialog standard).
            var dlg = new Microsoft.Win32.OpenFolderDialog
            {
                Title = "Choisir un dossier",
                InitialDirectory = !string.IsNullOrWhiteSpace(cheminInitial) && Directory.Exists(cheminInitial)
                    ? cheminInitial
                    : CheminsMetrologo.Racine
            };
            return dlg.ShowDialog() == true ? dlg.FolderName : null;
        }

        private static void TesterDossier(string libelle, string chemin, List<string> resultats)
        {
            if (string.IsNullOrWhiteSpace(chemin))
            {
                resultats.Add($"• {libelle} : (vide → défaut local) ✓");
                return;
            }
            try
            {
                Directory.CreateDirectory(chemin);
                string testFile = Path.Combine(chemin, $".test_{Guid.NewGuid():N}.tmp");
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);
                resultats.Add($"• {libelle} : {chemin}\n   → Lecture/écriture OK ✓");
            }
            catch (Exception ex)
            {
                resultats.Add($"• {libelle} : {chemin}\n   → ✖ ÉCHEC : {ex.Message}");
            }
        }
    }
}
