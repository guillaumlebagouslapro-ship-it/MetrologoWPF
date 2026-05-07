using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Journal;
using Metrologo.Views;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// VM de la fenêtre « Gérer les appareils » réservée à l'administrateur.
    /// Affiche la liste du catalogue local et permet d'éditer / supprimer chaque modèle.
    /// </summary>
    public partial class GestionAppareilsViewModel : ObservableObject
    {
        private readonly string _utilisateurActuel;

        public ObservableCollection<ModeleAppareil> Modeles => CatalogueAppareilsService.Instance.Modeles;

        [ObservableProperty] private string _cheminCatalogue = string.Empty;

        public bool CatalogueVide => Modeles.Count == 0;
        public int NbModeles => Modeles.Count;

        /// <summary>
        /// Vrai si l'utilisateur courant est administrateur. Contrôle la visibilité du
        /// bouton « Supprimer » dans l'UI : la modification reste libre pour tous, mais
        /// seul l'admin peut purger un modèle du catalogue (action irréversible).
        /// </summary>
        public bool EstAdmin { get; }

        public GestionAppareilsViewModel(string utilisateur, bool estAdmin = false)
        {
            _utilisateurActuel = utilisateur;
            EstAdmin = estAdmin;
            // Catalogue migré sur SQL Server (table dbo.T_CATALOGUE_APPAREILS).
            // L'ancien JSON local n'est plus utilisé — on affiche la source effective.
            CheminCatalogue = "SQL Server : Metrologo.dbo.T_CATALOGUE_APPAREILS";

            Modeles.CollectionChanged += (_, _) =>
            {
                OnPropertyChanged(nameof(CatalogueVide));
                OnPropertyChanged(nameof(NbModeles));
            };
        }

        [RelayCommand]
        private void ModifierModele(ModeleAppareil? modele)
        {
            if (modele == null) return;

            var vm = new EnregistrementAppareilViewModel(modele, _utilisateurActuel);
            var win = new EnregistrementAppareilWindow(vm) { Owner = Application.Current.MainWindow };
            win.ShowDialog();
        }

        [RelayCommand]
        private async Task SupprimerModeleAsync(ModeleAppareil? modele)
        {
            if (modele == null) return;

            // Garde-fou serveur : même si le bouton est censé être caché aux non-admins,
            // on refuse à nouveau ici (par exemple si quelqu'un déclenche la commande
            // par un autre chemin).
            if (!EstAdmin)
            {
                MessageBox.Show("Seul un administrateur peut supprimer un modèle du catalogue.",
                    "Action refusée", MessageBoxButton.OK, MessageBoxImage.Information);
                JournalLog.Warn(CategorieLog.Administration, "CATALOGUE_SUPPR_REFUSE",
                    $"Tentative de suppression du modèle « {modele.Nom} » par {_utilisateurActuel} (non admin).");
                return;
            }

            var confirm = MessageBox.Show(
                $"Supprimer définitivement le modèle « {modele.Nom } » du catalogue ?\n\n"
                + "Les appareils de ce modèle branchés sur le bus ne seront plus reconnus automatiquement.",
                "Confirmer la suppression",
                MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (confirm != MessageBoxResult.Yes) return;

            await CatalogueAppareilsService.Instance.SupprimerAsync(modele.Id);

            JournalLog.Warn(CategorieLog.Administration, "CATALOGUE_SUPPR",
                $"Modèle « {modele.Nom } » supprimé du catalogue par {_utilisateurActuel}.",
                new { modele.Id, modele.Nom });
        }

        /// <summary>
        /// Importe un (ou plusieurs) modèles d'appareil depuis un fichier .json — format
        /// hérité du catalogue local (cf. <c>AppareilsCatalogue.json</c>). Permet à l'admin
        /// d'ajouter en un clic un appareil pré-configuré (53131A, 53230A, SR620, etc.) sans
        /// ressaisir manuellement chaque commande SCPI dans la UI.
        /// </summary>
        [RelayCommand]
        private async Task ImporterJsonAsync()
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Importer un modèle d'appareil",
                Filter = "Fichiers JSON (*.json)|*.json|Tous les fichiers (*.*)|*.*",
                Multiselect = false
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                string json = await File.ReadAllTextAsync(dlg.FileName);
                int n = await CatalogueAppareilsService.Instance.ImporterDepuisJsonAsync(json);

                MessageBox.Show(
                    $"{n} modèle(s) importé(s) avec succès depuis {Path.GetFileName(dlg.FileName)}.",
                    "Import terminé", MessageBoxButton.OK, MessageBoxImage.Information);

                JournalLog.Info(CategorieLog.Administration, "CATALOGUE_IMPORT_UI",
                    $"{n} modèle(s) importé(s) depuis {dlg.FileName} par {_utilisateurActuel}.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Import échoué :\n\n{ex.Message}\n\n"
                  + "Vérifie que le fichier est un JSON valide au format catalogue.",
                    "Erreur d'import", MessageBoxButton.OK, MessageBoxImage.Error);

                JournalLog.Erreur(CategorieLog.Administration, "CATALOGUE_IMPORT_ERR",
                    $"Échec import JSON depuis {dlg.FileName} : {ex.Message}");
            }
        }

        [RelayCommand]
        private void OuvrirDossierCatalogue()
        {
            // Le catalogue est désormais en base SQL Server (T_CATALOGUE_APPAREILS) — il
            // n'y a plus de "dossier" à ouvrir. La commande est conservée pour ne pas
            // casser le binding XAML existant ; elle ne fait plus rien volontairement.
            try { /* no-op : SQL Server */ }
            catch { /* silencieux */ }
        }

        [RelayCommand]
        private void Fermer()
        {
            foreach (Window w in Application.Current.Windows)
            {
                if (w.DataContext == this) { w.Close(); return; }
            }
        }
    }
}
