using CommunityToolkit.Mvvm.ComponentModel;

namespace Metrologo.ViewModels
{
    public partial class AdminViewModel : ObservableObject
    {
        [ObservableProperty]
        private string _messageAdmin =
            "Bienvenue dans l'espace administrateur. Ici, on pourra ensuite ajouter la gestion des comptes, les paramètres machine, les droits d'accès et les réglages avancés.";
    }
}