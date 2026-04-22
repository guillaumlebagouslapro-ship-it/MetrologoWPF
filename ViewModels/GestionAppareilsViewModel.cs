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

        public GestionAppareilsViewModel(string utilisateur)
        {
            _utilisateurActuel = utilisateur;
            CheminCatalogue = CatalogueAppareilsService.Instance.CheminFichier;

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

        [RelayCommand]
        private void OuvrirDossierCatalogue()
        {
            try
            {
                var dir = Path.GetDirectoryName(CheminCatalogue);
                if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = dir,
                        UseShellExecute = true
                    });
                }
            }
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
