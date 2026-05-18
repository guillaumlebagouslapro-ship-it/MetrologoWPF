using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    /// <summary>
    /// Fenêtre d'aide statique expliquant la saisie des coefficients A et B
    /// (notation scientifique, séparateurs, ordres de grandeur typiques). Pure
    /// documentation — affichée à la demande depuis GestionModulesIncertitudeWindow.
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
