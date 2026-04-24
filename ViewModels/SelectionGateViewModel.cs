using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using System;
using System.Collections.Generic;

namespace Metrologo.ViewModels
{
    public partial class SelectionGateViewModel : ObservableObject
    {
        // Sentinelles pour IndexGate (conservées pour compat Excel / ancienne feuille de calcul).
        public const int Procedure10s = -2;
        public const int Procedure100s = -1;

        [ObservableProperty] private TypeMesure _typeMesure;
        [ObservableProperty] private int _gateSelectionneIndex = -1;
        [ObservableProperty] private bool _procedure100SChoisie;
        [ObservableProperty] private bool _procedure10SChoisie;
        [ObservableProperty] private string _messageErreur = string.Empty;

        public List<string> GateTimes { get; } = new()
        {
            "10 ms", "20 ms", "50 ms",
            "100 ms", "200 ms", "500 ms",
            "1 s", "2 s", "5 s",
            "10 s", "20 s", "50 s",
            "100 s", "200 s", "500 s", "1000 s"
        };

        public bool IsStabilite => TypeMesure == TypeMesure.Stabilite;
        public string Titre => IsStabilite ? "Mesure de stabilité" : "Mesure de fréquence";
        public bool HasError => !string.IsNullOrEmpty(MessageErreur);

        public int IndexGateResultat { get; private set; }
        public Action<bool>? CloseAction { get; set; }

        public SelectionGateViewModel(Mesure mesure)
        {
            TypeMesure = mesure.TypeMesure;

            // Valeur par défaut : gate 10 s (index 9) pour FreqAvantInterv, gate 1 s (index 6) sinon.
            if (TypeMesure == TypeMesure.FreqAvantInterv)
                GateSelectionneIndex = 9;
            else if (IsStabilite)
                Procedure10SChoisie = true;  // procédure auto par défaut en stabilité
            else
                GateSelectionneIndex = 6;
        }

        partial void OnGateSelectionneIndexChanged(int value)
        {
            if (value >= 0)
            {
                Procedure100SChoisie = false;
                Procedure10SChoisie = false;
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

        [RelayCommand]
        private void Valider()
        {
            if (GateSelectionneIndex >= 0)
            {
                IndexGateResultat = GateSelectionneIndex;
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
