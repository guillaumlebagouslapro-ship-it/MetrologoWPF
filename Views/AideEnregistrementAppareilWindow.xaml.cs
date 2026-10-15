using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    /// <summary>
    /// Fenêtre d'aide détaillant les réglages d'un appareil au catalogue (documentation statique,
    /// ni logique métier ni binding), ouverte depuis EnregistrementAppareilWindow.
    /// </summary>
    public partial class AideEnregistrementAppareilWindow : FluentWindow
    {
        public AideEnregistrementAppareilWindow()
        {
            InitializeComponent();
        }

        private void OnFermer(object sender, RoutedEventArgs e) => Close();
    }
}
