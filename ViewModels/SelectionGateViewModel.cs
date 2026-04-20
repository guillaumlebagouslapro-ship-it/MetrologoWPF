using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using System;
using System.Collections.Generic;

namespace Metrologo.ViewModels
{
    public partial class SelectionGateViewModel : ObservableObject
    {
        // Sentinelles pour IndexGate (conservées depuis l'ancien code pour compat Excel)
        public const int ProcedureEip = -3;
        public const int Procedure10s = -2;
        public const int Procedure100s = -1;

        [ObservableProperty] private TypeMesure _typeMesure;
        [ObservableProperty] private TypeAppareilIEEE _frequencemetre;
        [ObservableProperty] private int _gateSelectionneIndex = -1;
        [ObservableProperty] private bool _procedure100SChoisie;
        [ObservableProperty] private bool _procedure10SChoisie;
        [ObservableProperty] private bool _procedureEipChoisie;
        [ObservableProperty] private string _messageErreur = string.Empty;

        public List<string> GateTimes { get; } = new()
        {
            "10 ms", "20 ms", "50 ms",
            "100 ms", "200 ms", "500 ms",
            "1 s", "2 s", "5 s",
            "10 s", "20 s", "50 s", "100 s"
        };

        public bool IsStabilite => TypeMesure == TypeMesure.Stabilite;
        public bool IsEip => Frequencemetre == TypeAppareilIEEE.EIP;
        public bool AfficheProcedureEip => IsStabilite && IsEip;
        public bool AfficheProceduresStandard => IsStabilite && !IsEip;
        public string Titre => IsStabilite ? "Mesure de stabilité" : "Mesure de fréquence";
        public bool HasError => !string.IsNullOrEmpty(MessageErreur);

        public int IndexGateResultat { get; private set; }
        public Action<bool>? CloseAction { get; set; }

        public SelectionGateViewModel(Mesure mesure)
        {
            TypeMesure = mesure.TypeMesure;
            Frequencemetre = mesure.Frequencemetre;

            // Valeur par défaut : gate 1s (index 6) pour fréquence, 10s (index 9) pour FreqAvantInterv non-EIP
            if (TypeMesure == TypeMesure.FreqAvantInterv && !IsEip)
                GateSelectionneIndex = 9;
            else if (IsStabilite)
            {
                // Stabilité : pré-sélectionne une procédure selon l'instrument
                if (IsEip) ProcedureEipChoisie = true;
                else Procedure10SChoisie = true;
            }
            else
                GateSelectionneIndex = 6;
        }

        partial void OnGateSelectionneIndexChanged(int value)
        {
            if (value >= 0)
            {
                Procedure100SChoisie = false;
                Procedure10SChoisie = false;
                ProcedureEipChoisie = false;
            }
            OnPropertyChanged(nameof(HasError));
            MessageErreur = string.Empty;
        }

        partial void OnProcedure100SChoisieChanged(bool value)
        {
            if (value) { GateSelectionneIndex = -1; Procedure10SChoisie = false; }
        }

        partial void OnProcedure10SChoisieChanged(bool value)
        {
            if (value) { GateSelectionneIndex = -1; Procedure100SChoisie = false; }
        }

        partial void OnProcedureEipChoisieChanged(bool value)
        {
            if (value) GateSelectionneIndex = -1;
        }

        [RelayCommand]
        private void Valider()
        {
            if (GateSelectionneIndex >= 0)
            {
                IndexGateResultat = GateSelectionneIndex;
            }
            else if (ProcedureEipChoisie)
            {
                IndexGateResultat = ProcedureEip;
            }
            else if (Procedure100SChoisie)
            {
                IndexGateResultat = Procedure100s;
            }
            else if (Procedure10SChoisie)
            {
                IndexGateResultat = Procedure10s;
            }
            else
            {
                MessageErreur = "Sélectionnez un temps de porte ou une procédure.";
                OnPropertyChanged(nameof(HasError));
                return;
            }

            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
