using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Views;
using System;
using System.Collections.ObjectModel;
using System.Windows;

namespace Metrologo.ViewModels
{
    public partial class DispIncertViewModel : ObservableObject
    {
        [ObservableProperty] private ObservableCollection<Rubidium> _rubidiums = new();
        [ObservableProperty] private Rubidium? _rubidiumSelectionne;

        [ObservableProperty] private bool _editionDebloquee;
        [ObservableProperty] private ParametresIncertitudeGlobaux _parametresGlobaux = new();

        // Fréquence — une liste par appareil
        [ObservableProperty] private ObservableCollection<IncertitudeFrequence> _freqFixes = new();
        [ObservableProperty] private ObservableCollection<IncertitudeFrequence> _freqStanford = new();
        [ObservableProperty] private ObservableCollection<IncertitudeFrequence> _freqRacal = new();
        [ObservableProperty] private ObservableCollection<IncertitudeFrequence> _freqEip = new();

        // Stabilité
        [ObservableProperty] private ObservableCollection<IncertitudeStabilite> _stabFixes = new();
        [ObservableProperty] private ObservableCollection<IncertitudeStabilite> _stabStanford = new();
        [ObservableProperty] private ObservableCollection<IncertitudeStabilite> _stabRacal = new();

        // Autres mesures
        [ObservableProperty] private ObservableCollection<IncertitudeAutreMesure> _autresIntervalle = new();
        [ObservableProperty] private ObservableCollection<IncertitudeAutreMesure> _autresTachyContact = new();
        [ObservableProperty] private ObservableCollection<IncertitudeAutreMesure> _autresStroboscope = new();

        public Action<bool>? CloseAction { get; set; }

        public bool AucunRubidium => Rubidiums.Count == 0;

        public DispIncertViewModel()
        {
            Charger();
        }

        partial void OnRubidiumSelectionneChanged(Rubidium? value)
        {
            ChargerIncertitudesPourRubidium(value);
        }

        private void Charger()
        {
            // TODO : remplacer par requête SQL : S_REQ_RUBIDIUMS
            Rubidiums.Add(new Rubidium { Id = 1, Designation = "Rubidium A", FrequenceMoyenne = 10000000.0 });
            Rubidiums.Add(new Rubidium { Id = 2, Designation = "Rubidium B", FrequenceMoyenne = 10000000.0 });
            RubidiumSelectionne = Rubidiums[0];

            // TODO : remplacer par S_REQINFOS_INCERT
            ParametresGlobaux = new ParametresIncertitudeGlobaux
            {
                NbMesAccr = 30,
                TempsMesAccrSec = 10,
                IncertAccr = 1e-11
            };

            OnPropertyChanged(nameof(AucunRubidium));
        }

        private void ChargerIncertitudesPourRubidium(Rubidium? rubi)
        {
            // TODO : S_REQ_DISPINCERT_FREQ / S_REQ_DISPINCERT_STAB / S_REQ_DISPINCERT_AUTRES
            // Les coefficients seront chargés depuis SQL selon le rubidium sélectionné.
            FreqFixes.Clear();
            FreqStanford.Clear();
            FreqRacal.Clear();
            FreqEip.Clear();
            StabFixes.Clear();
            StabStanford.Clear();
            StabRacal.Clear();
            AutresIntervalle.Clear();
            AutresTachyContact.Clear();
            AutresStroboscope.Clear();
        }

        [RelayCommand]
        private void Modifier()
        {
            var vm = new MdpValidationViewModel(
                MdpValidationViewModel.MdpIncertitudes,
                titre: "Édition des incertitudes",
                sousTitre: "Ce mode de passe est requis pour modifier les coefficients.");

            var win = new MdpValidationWindow(vm) { Owner = Application.Current.MainWindow };
            if (win.ShowDialog() == true && vm.PasswordOk)
            {
                EditionDebloquee = true;
            }
        }

        [RelayCommand] private void Fermer() => CloseAction?.Invoke(true);
    }
}
