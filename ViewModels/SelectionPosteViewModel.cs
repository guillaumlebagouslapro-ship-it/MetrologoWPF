using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;

namespace Metrologo.ViewModels
{
    public partial class SelectionPosteViewModel : ObservableObject
    {
        public Action<bool>? OnPosteSelectionne { get; set; }

        /// <summary>
        /// Appelé lorsque l'utilisateur souhaite retourner à l'écran de connexion — par exemple
        /// s'il s'est trompé de compte au démarrage — sans devoir relancer l'application.
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