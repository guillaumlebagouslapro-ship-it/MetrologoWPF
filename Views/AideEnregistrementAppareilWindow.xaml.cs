using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    /// <summary>
    /// Petite fenêtre d'aide qui détaille tous les réglages d'un appareil au catalogue.
    /// Rien de dynamique ici : pas de logique métier, pas de binding. C'est juste de la
    /// doc qu'on affiche quand l'utilisateur la demande depuis EnregistrementAppareilWindow.
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
