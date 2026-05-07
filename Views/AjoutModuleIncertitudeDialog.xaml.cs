using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class AjoutModuleIncertitudeDialog : FluentWindow
    {
        public string NumModuleSaisi => TbNumModule.Text.Trim();
        public string NomAffichageSaisi => TbNomAffichage.Text.Trim();

        public AjoutModuleIncertitudeDialog()
        {
            InitializeComponent();
            Loaded += (_, _) => TbNumModule.Focus();
        }

        private void OnValider(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(NumModuleSaisi))
            {
                System.Windows.MessageBox.Show("Le numéro de module est requis.",
                    "Champ manquant",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                TbNumModule.Focus();
                return;
            }
            DialogResult = true;
        }

        private void OnAnnuler(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}
