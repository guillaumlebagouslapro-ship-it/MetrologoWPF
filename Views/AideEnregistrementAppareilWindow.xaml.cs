using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    /// <summary>
    /// Fenêtre d'aide statique expliquant tous les paramètres de configuration d'un appareil
    /// au catalogue. Pure information, aucune logique métier ni binding — c'est de la
    /// documentation utilisateur affichée à la demande depuis EnregistrementAppareilWindow.
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
