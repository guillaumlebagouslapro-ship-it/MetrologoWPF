using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// Étape réservée au poste Baie : choisir entre un scan GPIB automatique (les compteurs
    /// modernes qui répondent à *IDN?) ou des adresses fixes (les appareils legacy EIP / Racal /
    /// Stanford, qu'on désigne à la main). En Paillasse, on passe directement par-dessus : le scan
    /// est automatique. Structure calquée sur <see cref="SelectionPosteViewModel"/>.
    /// </summary>
    public partial class SelectionScanModeViewModel : ObservableObject
    {
        /// <summary>Appelé avec <c>true</c> si l'utilisateur choisit le scan GPIB, <c>false</c> pour les adresses fixes.</summary>
        public Action<bool>? OnModeChoisi { get; set; }

        /// <summary>Pour revenir en arrière, à la sélection du poste.</summary>
        public Action? OnRetour { get; set; }

        [RelayCommand]
        private void ChoisirScan() => OnModeChoisi?.Invoke(true);

        [RelayCommand]
        private void ChoisirAdressesFixes() => OnModeChoisi?.Invoke(false);

        [RelayCommand]
        private void Retour() => OnRetour?.Invoke();
    }
}
