using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;

namespace Metrologo.ViewModels
{
    public partial class MtxBusyViewModel : ObservableObject
    {
        [ObservableProperty] private string _titre = "Metrologo est déjà lancé";
        [ObservableProperty] private string _message =
            "Une autre instance de Metrologo est déjà en cours d'exécution sur cette machine. "
            + "Veuillez fermer l'instance existante avant de relancer l'application.";

        public bool AbandonDemande { get; private set; }
        public Action<bool>? CloseAction { get; set; }

        [RelayCommand]
        private void Abandonner()
        {
            AbandonDemande = true;
            CloseAction?.Invoke(false);
        }
    }
}
