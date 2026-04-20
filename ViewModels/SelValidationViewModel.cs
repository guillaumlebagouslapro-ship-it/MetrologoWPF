using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.ObjectModel;

namespace Metrologo.ViewModels
{
    public partial class SelValidationViewModel : ObservableObject
    {
        [ObservableProperty] private ObservableCollection<string> _validations = new();

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string? _validationSelectionnee;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasError))]
        private string _messageErreur = string.Empty;

        public bool HasError => !string.IsNullOrEmpty(MessageErreur);
        public string? ValidationChoisie { get; private set; }
        public Action<bool>? CloseAction { get; set; }

        public SelValidationViewModel()
        {
            // TODO : charger les validations depuis le classeur Excel ValidationMetrologo.xlsx
            // Parcourir AT_ZONESVALIDATION et lire les libellés dans les zones nommées
            Validations.Add("Correction de fréquence");
            Validations.Add("Valeur moyenne");
            Validations.Add("Écart-type / Incertitude globale");
            Validations.Add("Multiplicateur de fréquence");
            Validations.Add("Fréquences basses");
            Validations.Add("Fréquences hautes");
            Validations.Add("Incertitude sur la résolution");
            Validations.Add("Incertitude supplémentaire");
            Validations.Add("Incertitude résolution + supplémentaire");

            ValidationSelectionnee = Validations[0];
        }

        [RelayCommand]
        private void Valider()
        {
            if (string.IsNullOrEmpty(ValidationSelectionnee))
            {
                MessageErreur = "Sélectionnez un type de validation.";
                return;
            }

            ValidationChoisie = ValidationSelectionnee;
            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
