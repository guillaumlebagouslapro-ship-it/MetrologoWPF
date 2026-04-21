using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using System;
using System.Collections.ObjectModel;

namespace Metrologo.ViewModels
{
    public partial class ChoixRubidiumViewModel : ObservableObject
    {
        [ObservableProperty] private ObservableCollection<Rubidium> _rubidiums = new();
        [ObservableProperty] private Rubidium? _rubidiumSelectionne;
        [ObservableProperty] private bool _avecGPS;
        [ObservableProperty] private bool _sansGPS = true;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _messageErreur = string.Empty;

        public bool HasError => !string.IsNullOrEmpty(MessageErreur);

        public Rubidium? Resultat { get; private set; }
        public bool AvecGpsResultat { get; private set; }
        public Action<bool>? CloseAction { get; set; }

        public bool ListeVide => Rubidiums.Count == 0;

        public ChoixRubidiumViewModel()
        {
            ChargerRubidiums();
        }

        private void ChargerRubidiums()
        {
            // TODO : remplacer par une requête SQL : Select RUB_ID, RUB_ACTIF, RUB_DESIGNATION from TR_METROLOGO_RUBIDIUMS
            Rubidiums.Add(new Rubidium { Id = 1, Designation = "Rubidium A", FrequenceMoyenne = 10000000.0 });
            Rubidiums.Add(new Rubidium { Id = 2, Designation = "Rubidium B", FrequenceMoyenne = 10000000.0 });
            RubidiumSelectionne = Rubidiums[0];
            OnPropertyChanged(nameof(ListeVide));
        }

        partial void OnAvecGPSChanged(bool value)
        {
            if (value) SansGPS = false;
        }

        partial void OnSansGPSChanged(bool value)
        {
            if (value) AvecGPS = false;
        }

        [RelayCommand]
        private void Valider()
        {
            if (RubidiumSelectionne == null)
            {
                MessageErreur = "Veuillez sélectionner un rubidium.";
                return;
            }

            if (!AvecGPS && !SansGPS)
            {
                MessageErreur = "Veuillez spécifier le mode de raccordement.";
                return;
            }

            Resultat = RubidiumSelectionne;
            AvecGpsResultat = AvecGPS;
            Resultat.AvecGPS = AvecGPS;
            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
