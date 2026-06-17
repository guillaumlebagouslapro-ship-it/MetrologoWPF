using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    /// <summary>
    /// Fenêtre d'aide qui explique comment saisir les coefficients A et B : notation
    /// scientifique, séparateurs, ordres de grandeur courants. Juste de la doc, qu'on
    /// ouvre au besoin depuis GestionModulesIncertitudeWindow.
    /// </summary>
    public partial class AideSaisieIncertitudeWindow : FluentWindow
    {
        public AideSaisieIncertitudeWindow()
        {
            InitializeComponent();
        }

        private void OnFermer(object sender, RoutedEventArgs e) => Close();
    }
}
