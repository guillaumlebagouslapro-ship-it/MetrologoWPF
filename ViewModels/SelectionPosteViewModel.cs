using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;

namespace Metrologo.ViewModels
{
    public partial class SelectionPosteViewModel : ObservableObject
    {
        public Action<bool>? OnPosteSelectionne { get; set; }

        /// <summary>
        /// Invoqué quand l'utilisateur veut revenir à l'écran de connexion (ex. mauvais
        /// utilisateur sélectionné au démarrage) sans avoir à redémarrer l'application.
        /// </summary>
        public Action? OnRetour { get; set; }

        [RelayCommand]
        private void ChoisirBaie() => OnPosteSelectionne?.Invoke(true);

        [RelayCommand]
        private void ChoisirPaillasse() => OnPosteSelectionne?.Invoke(false);

        [RelayCommand]
        private void Retour() => OnRetour?.Invoke();
    }
}