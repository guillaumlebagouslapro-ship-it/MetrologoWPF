using System.Windows;
using Metrologo.Models;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class AjoutModuleIncertitudeDialog : FluentWindow
    {
        public string NumModuleSaisi => TbNumModule.Text.Trim();
        public string NomAffichageSaisi => TbNomAffichage.Text.Trim();
        public bool SansTempsDeMesure => CbSansTemps.IsChecked == true;

        public AjoutModuleIncertitudeDialog() : this(null) { }

        /// <summary>
        /// <paramref name="categorie"/> = le type de mesure en cours. Pour Tachy/Strobo on
        /// coche d'office « Sans temps de mesure » (c'est le cas courant) ; sinon on laisse
        /// la case vide.
        /// </summary>
        public AjoutModuleIncertitudeDialog(TypeMesure? categorie)
        {
            InitializeComponent();
            if (categorie == TypeMesure.TachyContact
                || categorie == TypeMesure.TachyOptique
                || categorie == TypeMesure.Stroboscope)
                CbSansTemps.IsChecked = true;
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
