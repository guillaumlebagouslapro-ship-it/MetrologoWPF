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

        public DispIncertViewModel()
        {
            ChargerDonneesSimulees();
        }

        partial void OnRubidiumSelectionneChanged(Rubidium? value)
        {
            // TODO : recharger les données depuis SQL selon le rubidium sélectionné
            ChargerIncertitudesPourRubidium(value);
        }

        private void ChargerDonneesSimulees()
        {
            // TODO : remplacer par requête SQL : S_REQ_RUBIDIUMS
            Rubidiums.Add(new Rubidium { Id = 1, Designation = "FS725 - SN 12345", FrequenceMoyenne = 10000000.0 });
            Rubidiums.Add(new Rubidium { Id = 2, Designation = "FS725 - SN 67890", FrequenceMoyenne = 10000000.0 });
            Rubidiums.Add(new Rubidium { Id = 3, Designation = "LPRO-101 - SN 54321", FrequenceMoyenne = 10000000.0 });
            RubidiumSelectionne = Rubidiums[0];

            // TODO : remplacer par S_REQINFOS_INCERT
            ParametresGlobaux = new ParametresIncertitudeGlobaux
            {
                NbMesAccr = 30,
                TempsMesAccrSec = 10,
                IncertAccr = 1e-11
            };

            ChargerAutresMesures();
        }

        private void ChargerIncertitudesPourRubidium(Rubidium? rubi)
        {
            if (rubi == null) return;

            // TODO : S_REQ_DISPINCERT_FREQ avec rubi.Id et l'appareil
            FreqFixes = new ObservableCollection<IncertitudeFrequence>
            {
                new() { Plage = "10 Hz – 100 kHz",    Raccord = "Allouis", CoeffA = 1.2e-10, CoeffB = 5e-13 },
                new() { Plage = "100 kHz – 10 MHz",   Raccord = "Allouis", CoeffA = 8.0e-11, CoeffB = 3e-13 },
                new() { Plage = "10 Hz – 100 kHz",    Raccord = "GPS",     CoeffA = 1.0e-10, CoeffB = 4e-13 },
                new() { Plage = "100 kHz – 10 MHz",   Raccord = "GPS",     CoeffA = 7.0e-11, CoeffB = 2e-13 },
            };

            FreqStanford = new ObservableCollection<IncertitudeFrequence>
            {
                new() { Plage = "DC – 200 MHz (1MΩ)", Raccord = "Allouis", CoeffA = 5.0e-11, CoeffB = 2e-13 },
                new() { Plage = "DC – 200 MHz (50Ω)", Raccord = "Allouis", CoeffA = 4.5e-11, CoeffB = 2e-13 },
                new() { Plage = "UHF (C)",             Raccord = "GPS",     CoeffA = 6.0e-11, CoeffB = 3e-13 },
            };

            FreqRacal = new ObservableCollection<IncertitudeFrequence>
            {
                new() { Plage = "Entrée A (50Ω)",  Raccord = "Allouis", CoeffA = 1.5e-10, CoeffB = 5e-13 },
                new() { Plage = "Entrée A (1MΩ)",  Raccord = "Allouis", CoeffA = 1.5e-10, CoeffB = 5e-13 },
                new() { Plage = "Entrée C",         Raccord = "GPS",     CoeffA = 2.0e-10, CoeffB = 7e-13 },
            };

            FreqEip = new ObservableCollection<IncertitudeFrequence>
            {
                new() { Plage = "Bande 1 (B1)", Raccord = "Allouis", CoeffA = 3.0e-10, CoeffB = 1e-12 },
                new() { Plage = "Bande 2 (B2)", Raccord = "Allouis", CoeffA = 3.5e-10, CoeffB = 1e-12 },
                new() { Plage = "Bande 3 (B3)", Raccord = "GPS",     CoeffA = 4.0e-10, CoeffB = 1.2e-12 },
            };

            // TODO : S_REQ_DISPINCERT_STAB
            StabFixes = new ObservableCollection<IncertitudeStabilite>
            {
                new() { TempsDeMesure = "1 s",   Valeur = 2e-11 },
                new() { TempsDeMesure = "10 s",  Valeur = 7e-12 },
                new() { TempsDeMesure = "100 s", Valeur = 3e-12 },
            };

            StabStanford = new ObservableCollection<IncertitudeStabilite>
            {
                new() { TempsDeMesure = "10 ms", Valeur = 1e-9 },
                new() { TempsDeMesure = "1 s",   Valeur = 1e-11 },
                new() { TempsDeMesure = "100 s", Valeur = 2e-12 },
            };

            StabRacal = new ObservableCollection<IncertitudeStabilite>
            {
                new() { TempsDeMesure = "10 ms", Valeur = 5e-9 },
                new() { TempsDeMesure = "1 s",   Valeur = 5e-11 },
                new() { TempsDeMesure = "10 s",  Valeur = 1e-11 },
            };
        }

        private void ChargerAutresMesures()
        {
            // TODO : S_REQ_DISPINCERT_AUTRES
            AutresIntervalle = new ObservableCollection<IncertitudeAutreMesure>
            {
                new() { Libelle = "Intervalle de temps (défaut)", CoeffA = 1e-8, CoeffB = 1e-10 },
            };
            AutresTachyContact = new ObservableCollection<IncertitudeAutreMesure>
            {
                new() { Libelle = "Tachymétrie par contacts",  CoeffA = 1e-4, CoeffB = 1e-6 },
            };
            AutresStroboscope = new ObservableCollection<IncertitudeAutreMesure>
            {
                new() { Libelle = "Stroboscope",                CoeffA = 5e-4, CoeffB = 5e-6 },
            };
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
