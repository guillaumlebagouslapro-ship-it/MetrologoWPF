using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// Étape (poste Baie uniquement) : choisir entre un scan GPIB automatique (compteurs modernes
    /// répondant à *IDN?) ou des adresses fixes (appareils legacy EIP / Racal / Stanford qu'on
    /// sélectionne manuellement). En Paillasse, cette étape est sautée (scan automatique).
    /// Calqué sur <see cref="SelectionPosteViewModel"/>.
    /// </summary>
    public partial class SelectionScanModeViewModel : ObservableObject
    {
        /// <summary>Invoqué avec <c>true</c> pour le scan GPIB, <c>false</c> pour les adresses fixes.</summary>
        public Action<bool>? OnModeChoisi { get; set; }

        /// <summary>Retour à la sélection du poste.</summary>
        public Action? OnRetour { get; set; }

        [RelayCommand]
        private void ChoisirScan() => OnModeChoisi?.Invoke(true);

        [RelayCommand]
        private void ChoisirAdressesFixes() => OnModeChoisi?.Invoke(false);

        [RelayCommand]
        private void Retour() => OnRetour?.Invoke();
    }
}
