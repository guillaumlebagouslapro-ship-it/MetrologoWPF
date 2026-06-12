using System.Windows;

namespace Metrologo.Services
{
    /// <summary>
    /// MessageBox garanti AU-DESSUS de toutes les fenêtres, y compris celles des autres
    /// applications. Cas d'usage : en fin de mesure, le rapport Excel est au premier
    /// plan — un MessageBox classique apparaîtrait DERRIÈRE lui et l'utilisateur ne
    /// verrait pas que la mesure a échoué (ex. valeur hors du module d'incertitude).
    /// Astuce : le MessageBox est rattaché à une fenêtre propriétaire invisible et
    /// Topmost, ce qui le rend lui-même topmost — il passe donc par-dessus Excel.
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
                // Filet de sécurité : si la fenêtre propriétaire ne peut pas être créée,
                // on retombe sur le MessageBox classique (possiblement derrière Excel,
                // mais le message n'est pas perdu).
                return MessageBox.Show(message, titre, boutons, icone);
            }
            finally
            {
                try { proprietaire?.Close(); } catch { /* best-effort */ }
            }
        }
    }
}
