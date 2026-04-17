using Metrologo.Services;
using Metrologo.ViewModels;
using Metrologo.Views;
using System.Windows;

namespace Metrologo
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // 1. On prépare la fenêtre de Login
            var authService = new AuthService();
            var loginVM = new LoginViewModel(authService);
            var loginWin = new LoginWindow(loginVM);

            // 2. On l'affiche en mode "Bloquant" (Dialog)
            if (loginWin.ShowDialog() == true)
            {
                // 3. Si l'utilisateur a réussi à se connecter, on lance le vrai logiciel !
                var mainWin = new MainWindow();

                // On transmet l'utilisateur connecté au MainViewModel (si vous avez ajouté la propriété)
                if (mainWin.DataContext is MainViewModel mainVM)
                {
                    mainVM.UtilisateurConnecte = loginVM.UtilisateurSession;
                }

                mainWin.Show();
            }
            else
            {
                // Si l'utilisateur ferme la fenêtre de login ou clique sur Quitter, on arrête tout
                Shutdown();
            }
        }
    }
}