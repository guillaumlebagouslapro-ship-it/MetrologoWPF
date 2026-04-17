using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using System.Windows;

namespace Metrologo.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(EstAdmin))]
        [NotifyPropertyChangedFor(nameof(TexteUtilisateurConnecte))]
        private Utilisateur? _utilisateurConnecte;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(EstEnModeAdmin))]
        [NotifyPropertyChangedFor(nameof(TexteMode))]
        private object? _vueActuelle;

        [ObservableProperty]
        private string _titreApplication = "Metrologo v2026 - Migration Delphi";

        private readonly AccueilViewModel _accueilViewModel = new();
        private readonly AdminViewModel _adminViewModel = new();

        public bool EstAdmin => UtilisateurConnecte?.Role == RoleUtilisateur.Administrateur;
        public bool EstEnModeAdmin => VueActuelle is AdminViewModel;
        public string TexteMode => EstEnModeAdmin ? "Mode : Administration" : "Mode : Exploitation";
        public string TexteUtilisateurConnecte =>
            UtilisateurConnecte == null
                ? "Utilisateur : non connecté"
                : $"Utilisateur : {UtilisateurConnecte.Login} ({UtilisateurConnecte.Role})";
        public string RubidiumActifTexte => _accueilViewModel.RubidiumActifTexte;

        public MainViewModel()
        {
            VueActuelle = _accueilViewModel;
            _accueilViewModel.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(AccueilViewModel.RubidiumActifTexte))
                    OnPropertyChanged(nameof(RubidiumActifTexte));
            };
        }

        partial void OnUtilisateurConnecteChanged(Utilisateur? value)
        {
            OnPropertyChanged(nameof(EstAdmin));
            OnPropertyChanged(nameof(TexteUtilisateurConnecte));
        }

        partial void OnVueActuelleChanged(object? value)
        {
            OnPropertyChanged(nameof(EstEnModeAdmin));
            OnPropertyChanged(nameof(TexteMode));
        }

        [RelayCommand]
        private void AllerAccueil() => VueActuelle = _accueilViewModel;

        [RelayCommand]
        private void OuvrirAdmin()
        {
            if (!EstAdmin) return;
            VueActuelle = _adminViewModel;
        }

        [RelayCommand]
        private void RetourAccueil() => VueActuelle = _accueilViewModel;

        [RelayCommand]
        private void Quitter() => Application.Current.Shutdown();
    }
}