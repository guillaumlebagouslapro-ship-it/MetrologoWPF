using System.Windows;

namespace Metrologo.Services
{
    /// <summary>
    /// MessageBox garanti au-dessus de toutes les fenêtres, y compris Excel. Un
    /// MessageBox classique passerait derrière le rapport ouvert en fin de mesure
    /// (ex. valeur hors module d'incertitude). Astuce : fenêtre propriétaire invisible
    /// et Topmost → le MessageBox hérite du Topmost.
    /// </summary>
    public static class MessageBoxPremierPlan
    {
        public static MessageBoxResult Afficher(string message, string titre,
            MessageBoxButton boutons = MessageBoxButton.OK,
            MessageBoxImage icone = MessageBoxImage.None)
        {
            Window? proprietaire = null;
            try
            {
                proprietaire = new Window
                {
                    WindowStyle = WindowStyle.None,
                    ResizeMode = ResizeMode.NoResize,
                    AllowsTransparency = true,
                    Opacity = 0,
                    ShowInTaskbar = false,
                    Topmost = true,
                    Width = 1,
                    Height = 1,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                };
                proprietaire.Show();
                proprietaire.Activate();
                return MessageBox.Show(proprietaire, message, titre, boutons, icone);
            }
            catch
            {
                // Fallback : MessageBox classique (possiblement derrière Excel, mais pas perdu).
                return MessageBox.Show(message, titre, boutons, icone);
            }
            finally
            {
                try { proprietaire?.Close(); } catch { /* best-effort */ }
            }
        }
    }
}
