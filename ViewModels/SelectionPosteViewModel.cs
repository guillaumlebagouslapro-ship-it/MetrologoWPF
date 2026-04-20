using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;

namespace Metrologo.ViewModels
{
    public partial class SelectionPosteViewModel : ObservableObject
    {
        public Action<bool>? OnPosteSelectionne { get; set; }

        [RelayCommand]
        private void ChoisirBaie() => OnPosteSelectionne?.Invoke(true);

        [RelayCommand]
        private void ChoisirPaillasse() => OnPosteSelectionne?.Invoke(false);
    }
}