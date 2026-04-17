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


        // Menus déroulants
        public IEnumerable<TypeAppareilIEEE> Appareils => Enum.GetValues(typeof(TypeAppareilIEEE)).Cast<TypeAppareilIEEE>();
        public IEnumerable<TypeMesure> TypesMesure => Enum.GetValues(typeof(TypeMesure)).Cast<TypeMesure>();

        // Options extraites de vos fichiers .ini
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

        // Propriétés de visibilité (Calculées en temps réel)
        public List<int> MeasurementCounts => Enumerable.Range(1, 100).ToList();
        public bool IsStanford => MesureConfig.Frequencemetre == TypeAppareilIEEE.Stanford;
        public bool IsRacal => MesureConfig.Frequencemetre == TypeAppareilIEEE.Racal;
        public bool IsEip => MesureConfig.Frequencemetre == TypeAppareilIEEE.EIP;

        public bool ShowGateSettings =>
            MesureConfig.TypeMesure == TypeMesure.Frequence ||
            MesureConfig.TypeMesure == TypeMesure.Stabilite;

        public bool ShowCoupling => (IsStanford && MesureConfig.InputIndex != 2) ||
                                    (IsRacal && MesureConfig.InputIndex != 2);

        // Gestion du Mode de calcul (Fixé pour éviter l'erreur TwoWay)
        public bool IsModeDirect
        {
            get => MesureConfig.ModeMesure == ModeMesure.Direct;
            set { if (value) { MesureConfig.ModeMesure = ModeMesure.Direct; RefreshAll(); } }
        }

        public bool IsModeIndirect
        {
            get => MesureConfig.ModeMesure == ModeMesure.Indirect;
            set { if (value) { MesureConfig.ModeMesure = ModeMesure.Indirect; RefreshAll(); } }
        }

        // Rafraîchissement global
        public void RefreshAll()
        {
            OnPropertyChanged(nameof(IsStanford));
            OnPropertyChanged(nameof(IsRacal));
            OnPropertyChanged(nameof(IsEip));
            OnPropertyChanged(nameof(IsModeDirect));
            OnPropertyChanged(nameof(IsModeIndirect));
            OnPropertyChanged(nameof(ShowGateSettings));
            OnPropertyChanged(nameof(ShowCoupling));
        }

        public void OnTypeMesureChanged()
        {
            if (MesureConfig.TypeMesure == TypeMesure.Interval)
            {
                MesureConfig.NbMesures = 1;
            }
            else
            {
                // On remet à 30 si on quitte le mode intervalle
                MesureConfig.NbMesures = 30;
            }

            // On prévient l'interface que la valeur a changé
            OnPropertyChanged(nameof(MesureConfig));
            RefreshAll();
        }

        // Actions de fermeture
        public Action<bool>? CloseAction { get; set; }
        [RelayCommand] private void Valider() => CloseAction?.Invoke(true);
        [RelayCommand] private void Annuler() => CloseAction?.Invoke(false);
    }
}