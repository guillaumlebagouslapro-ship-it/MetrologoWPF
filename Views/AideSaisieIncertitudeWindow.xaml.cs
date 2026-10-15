using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    /// <summary>
    /// Fenêtre d'aide sur la saisie des coefficients A et B (notation scientifique, séparateurs,
    /// ordres de grandeur) : documentation statique, ouverte depuis GestionModulesIncertitudeWindow.
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
