using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Services.Journal;

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

            var utilisateur = await _authService.AuthentifierAsync(LoginStr, motDePasse);

            // ---------------------------------------------------------------------
            // TODO BDD-CENTRALE : retirer ce fallback une fois SQL Server Express
            // installé et la table T_UTILISATEURS peuplée en prod. Tant que le
            // serveur n'est pas joignable, AuthentifierAsync retourne null et on
            // tombe ici — sans ce bloc, plus aucune connexion possible.
            // Le compte admin par défaut côté serveur (admin/admin123) prendra
            // le relais une fois la base opérationnelle.
            // ---------------------------------------------------------------------
            if (utilisateur == null)
            {
                if ((LoginStr == "admin" && motDePasse == "admin") ||
                    (LoginStr == "metrologo" && motDePasse == "metrologo") ||
                    (LoginStr == "test" && motDePasse == "test"))
                {
                    utilisateur = new Utilisateur
                    {
                        Login = LoginStr,
                        Role = LoginStr == "admin" ? RoleUtilisateur.Administrateur : RoleUtilisateur.Utilisateur
                    };
                }
            }

            if (utilisateur != null)
            {
                UtilisateurSession = utilisateur;
                CloseAction?.Invoke(true);
            }
            else
            {
                MessageErreur = "Identifiant ou mot de passe incorrect.";
                // Tentative infructueuse — loggée hors session (sera rattachée à la prochaine session si elle s'ouvre)
                Journal.Warn(CategorieLog.Authentification, "ECHEC_CONNEXION",
                    $"Tentative de connexion échouée pour « {LoginStr} ».");
            }
        }

        [RelayCommand]
        private void Quitter() => CloseAction?.Invoke(false);
    }
}