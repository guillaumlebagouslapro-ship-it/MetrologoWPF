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
        private string _titreApplication = "Metrologo v2026";

        [ObservableProperty]
        private bool _estSurBaie;

        private readonly AccueilViewModel _accueilViewModel = new();
        private readonly AdminViewModel _adminViewModel = new();
        private readonly SelectionPosteViewModel _selectionPosteViewModel = new();

        public bool EstAdmin => UtilisateurConnecte?.Role == RoleUtilisateur.Administrateur;
        public bool EstEnModeAdmin => VueActuelle is AdminViewModel;
        public bool EstEnSelectionPoste => VueActuelle is SelectionPosteViewModel;
        public string TexteMode => EstEnModeAdmin ? "Mode : Administration" : "Mode : Exploitation";
        public string TexteUtilisateurConnecte =>
            UtilisateurConnecte == null
                ? "Utilisateur : non connecté"
                : $"Utilisateur : {UtilisateurConnecte.Login} ({(EstSurBaie ? "Baie" : "Paillasse")})";
        public string RubidiumActifTexte => _accueilViewModel.RubidiumActifTexte;

        public MainViewModel()
        {
            // On commence par la sélection du poste
            VueActuelle = _selectionPosteViewModel;

            // Quand l'utilisateur choisit son poste
            _selectionPosteViewModel.OnPosteSelectionne = (choixBaie) =>
            {
                EstSurBaie = choixBaie;
                _accueilViewModel.EstSurBaie = choixBaie;
                OnPropertyChanged(nameof(TexteUtilisateurConnecte));
                VueActuelle = _accueilViewModel;
            };

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
            OnPropertyChanged(nameof(EstEnSelectionPoste));
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