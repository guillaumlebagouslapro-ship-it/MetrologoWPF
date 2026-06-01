using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services.Journal;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// CRUD du catalogue des rubidiums (Syref, Redondances, etc.). Édition inline du
    /// nom et de la fréquence moyenne ; ajout / suppression libres. La persistance
    /// se fait dans <see cref="Preferences"/> au clic sur « Enregistrer ».
    /// </summary>
    public partial class GestionCatalogueRubidiumsViewModel : ObservableObject
    {
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(CatalogueVide))]
        private ObservableCollection<RubidiumEditable> _rubidiums = new();

        [ObservableProperty] private RubidiumEditable? _selectionne;

        public bool CatalogueVide => Rubidiums.Count == 0;
        public Action<bool>? CloseAction { get; set; }

        public GestionCatalogueRubidiumsViewModel()
        {
            foreach (var r in Preferences.CatalogueRubidiums)
            {
                Rubidiums.Add(new RubidiumEditable
                {
                    Id = r.Id,
                    Designation = r.Designation,
                    FrequenceMoyenne = r.FrequenceMoyenne,
                });
            }
            Rubidiums.CollectionChanged += (_, __) => OnPropertyChanged(nameof(CatalogueVide));
            Selectionne = Rubidiums.FirstOrDefault();
        }

        [RelayCommand]
        private void Ajouter()
        {
            int nextId = Rubidiums.Count == 0 ? 1 : Rubidiums.Max(r => r.Id) + 1;
            var nouveau = new RubidiumEditable
            {
                Id = nextId,
                Designation = $"Référence {nextId}",
                FrequenceMoyenne = 10_000_000.0,
            };
            Rubidiums.Add(nouveau);
            Selectionne = nouveau;
        }

        [RelayCommand]
        private void Supprimer(RubidiumEditable? item)
        {
            var cible = item ?? Selectionne;
            if (cible == null) return;

            var conf = MessageBox.Show(
                $"Supprimer « {cible.Designation} » du catalogue ?",
                "Confirmation",
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (conf != MessageBoxResult.Yes) return;

            Rubidiums.Remove(cible);
            Selectionne = Rubidiums.FirstOrDefault();
        }

        [RelayCommand]
        private void Enregistrer()
        {
            // Validation : pas de désignation vide, pas de fréquence ≤ 0.
            foreach (var r in Rubidiums)
            {
                if (string.IsNullOrWhiteSpace(r.Designation))
                {
                    MessageBox.Show("Une désignation est vide. Corrige avant d'enregistrer.",
                        "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (r.FrequenceMoyenne <= 0)
                {
                    MessageBox.Show($"Fréquence invalide pour « {r.Designation} ». Doit être > 0 Hz.",
                        "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            var snapshot = Rubidiums.Select(r => new Rubidium
            {
                Id = r.Id,
                Designation = r.Designation.Trim(),
                FrequenceMoyenne = r.FrequenceMoyenne,
            }).ToList();

            Preferences.SauvegarderCatalogueRubidiums(snapshot);

            // Si le rubidium actif fait partie du catalogue, on rafraîchit ses infos
            // (au cas où l'admin aurait renommé ou modifié la fréquence d'une entrée
            // actuellement sélectionnée).
            var actif = EtatApplication.RubidiumActif;
            if (actif != null && !actif.EstReglageManuel)
            {
                var maj = snapshot.FirstOrDefault(r => r.Id == actif.Id);
                if (maj != null)
                {
                    actif.Designation = maj.Designation;
                    actif.FrequenceMoyenne = maj.FrequenceMoyenne;
                    EtatApplication.NotifierRubidiumActifChange();
                }
                else
                {
                    // L'entrée a été supprimée du catalogue : on retombe sur le 1er
                    // rubidium du catalogue (sinon null), pour éviter d'avoir un
                    // rubidium actif fantôme qui n'existe plus.
                    EtatApplication.RubidiumActif = snapshot.FirstOrDefault();
                }
            }

            Journal.Info(CategorieLog.Rubidium, "CATALOGUE_RUBIDIUMS_MAJ",
                $"Catalogue mis à jour ({snapshot.Count} entrée(s))");

            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }

    /// <summary>
    /// Wrapper observable autour de <see cref="Rubidium"/> pour permettre l'édition
    /// inline dans la grille (TwoWay binding sur Designation et FrequenceMoyenne).
    /// </summary>
    public partial class RubidiumEditable : ObservableObject
    {
        [ObservableProperty] private int _id;
        [ObservableProperty] private string _designation = string.Empty;
        [ObservableProperty] private double _frequenceMoyenne;
    }
}
