using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.Generic;
using System.Linq;
using Metrologo.Models;

namespace Metrologo.ViewModels
{
    public partial class ConfigurationViewModel : ObservableObject
    {
        [ObservableProperty]
        private Mesure _mesureConfig = new Mesure();

        [ObservableProperty]
        private bool _estSurBaie = true;

        // Menus déroulants
        public IEnumerable<TypeAppareilIEEE> Appareils => Enum.GetValues(typeof(TypeAppareilIEEE)).Cast<TypeAppareilIEEE>();
        public IEnumerable<TypeMesure> TypesMesure => Enum.GetValues(typeof(TypeMesure)).Cast<TypeMesure>();

        public List<string> StanfordRanges => new() { "1 MΩ (A)", "50 Ω (A)", "UHF (C)" };
        public List<string> RacalRanges => new() { "Entrée A", "Entrée B", "Entrée C" };
        public List<string> EipRanges => new() { "Bande 1", "Bande 2", "Bande 3" };
        public List<string> Couplings => new() { "AC", "DC" };
        public List<string> GateTimes => new()
        {
            "10 ms", "20 ms", "50 ms",
            "100 ms", "200 ms", "500 ms",
            "1 s", "2 s", "5 s",
            "10 s", "20 s", "50 s", "100 s"
        };

        public List<int> MeasurementCounts => Enumerable.Range(1, 100).ToList();

        public bool IsStanford => MesureConfig.Frequencemetre == TypeAppareilIEEE.Stanford;
        public bool IsRacal => MesureConfig.Frequencemetre == TypeAppareilIEEE.Racal;
        public bool IsEip => MesureConfig.Frequencemetre == TypeAppareilIEEE.EIP;

        public bool ShowGateSettings =>
            MesureConfig.TypeMesure == TypeMesure.Frequence ||
            MesureConfig.TypeMesure == TypeMesure.Stabilite;

        public bool ShowCoupling => (IsStanford && MesureConfig.InputIndex != 2) ||
                                    (IsRacal && MesureConfig.InputIndex != 2);

        // Source du signal : visible seulement pour le type "Fréquence"
        public bool ShowSourceMesure => MesureConfig.TypeMesure == TypeMesure.Frequence;

        public bool IsSourceFrequencemetre
        {
            get => MesureConfig.SourceMesure == SourceMesure.Frequencemetre;
            set { if (value) { MesureConfig.SourceMesure = SourceMesure.Frequencemetre; RefreshAll(); } }
        }

        public bool IsSourceGenerateur
        {
            get => MesureConfig.SourceMesure == SourceMesure.Generateur;
            set { if (value) { MesureConfig.SourceMesure = SourceMesure.Generateur; RefreshAll(); } }
        }

        // Indirect disponible : pas en paillasse, pas EIP, pas intervalle / tachy / stroboscope
        public bool IndirectDisponible =>
            EstSurBaie
            && !IsEip
            && MesureConfig.TypeMesure != TypeMesure.Interval
            && MesureConfig.TypeMesure != TypeMesure.TachyContact
            && MesureConfig.TypeMesure != TypeMesure.Stroboscope;

        public bool IsModeDirect
        {
            get => MesureConfig.ModeMesure == ModeMesure.Direct;
            set { if (value) { MesureConfig.ModeMesure = ModeMesure.Direct; RefreshAll(); } }
        }

        public bool IsModeIndirect
        {
            get => MesureConfig.ModeMesure == ModeMesure.Indirect;
            set
            {
                if (value && !IndirectDisponible) return;
                if (value) { MesureConfig.ModeMesure = ModeMesure.Indirect; RefreshAll(); }
            }
        }

        public void RefreshAll()
        {
            OnPropertyChanged(nameof(IsStanford));
            OnPropertyChanged(nameof(IsRacal));
            OnPropertyChanged(nameof(IsEip));
            OnPropertyChanged(nameof(IsModeDirect));
            OnPropertyChanged(nameof(IsModeIndirect));
            OnPropertyChanged(nameof(ShowGateSettings));
            OnPropertyChanged(nameof(ShowCoupling));
            OnPropertyChanged(nameof(ShowSourceMesure));
            OnPropertyChanged(nameof(IsSourceFrequencemetre));
            OnPropertyChanged(nameof(IsSourceGenerateur));
            OnPropertyChanged(nameof(IndirectDisponible));
        }

        public void OnTypeMesureChanged()
        {
            if (MesureConfig.TypeMesure == TypeMesure.Interval)
                MesureConfig.NbMesures = 1;
            else
                MesureConfig.NbMesures = 30;

            // Si on quitte le type Fréquence, on repasse sur Fréquencemètre par défaut
            if (MesureConfig.TypeMesure != TypeMesure.Frequence)
                MesureConfig.SourceMesure = SourceMesure.Frequencemetre;

            // Si le mode indirect n'est plus dispo, on bascule sur Direct
            if (MesureConfig.ModeMesure == ModeMesure.Indirect && !IndirectDisponible)
                MesureConfig.ModeMesure = ModeMesure.Direct;

            OnPropertyChanged(nameof(MesureConfig));
            RefreshAll();
        }

        public Action<bool>? CloseAction { get; set; }
        [RelayCommand] private void Valider() => CloseAction?.Invoke(true);
        [RelayCommand] private void Annuler() => CloseAction?.Invoke(false);
    }
}
