using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services;

namespace Metrologo.ViewModels
{
    // LE MOT-CLÉ 'partial' EST OBLIGATOIRE ICI
    public partial class LoginViewModel : ObservableObject
    {
        private readonly IAuthService _authService;

        [ObservableProperty]
        private string _loginStr = string.Empty; // Générera automatiquement 'LoginStr'

        [ObservableProperty]
        private string _messageErreur = string.Empty; // Générera automatiquement 'MessageErreur'

        public Utilisateur? UtilisateurSession { get; private set; }
        public Action<bool>? CloseAction { get; set; }

        public LoginViewModel(IAuthService authService)
        {
            _authService = authService;
        }

        public async Task SeConnecterAsync(string motDePasse)
        {
            MessageErreur = "Connexion en cours..."; // Maintenant accessible

            var utilisateur = await _authService.AuthentifierAsync(LoginStr, motDePasse); // Maintenant accessible

            if (utilisateur != null)
            {
                UtilisateurSession = utilisateur;
                CloseAction?.Invoke(true);
            }
            else
            {
                MessageErreur = "Identifiant ou mot de passe incorrect.";
            }
        }

        [RelayCommand]
        private void Quitter() => CloseAction?.Invoke(false);
    }
}