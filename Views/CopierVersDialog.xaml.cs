using System.Linq;
using System.Windows;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    /// <summary>
    /// Dialog minimal pour choisir la catégorie cible quand on copie un module
    /// d'incertitude d'un sous-dossier vers un autre. Exclut la catégorie source
    /// pour éviter une copie no-op.
    /// </summary>
    public partial class CopierVersDialog : FluentWindow
    {
        public TypeMesure CategorieChoisie { get; private set; }

        public CopierVersDialog(TypeMesure categorieSource, string numModule)
        {
            InitializeComponent();

            string libelleSource = EnTetesMesureHelper.LibelleType(categorieSource);
            TbInfo.Text = $"Le module « {numModule } » est actuellement dans « {libelleSource} ».\n"
                + "Choisis la catégorie cible. Si un module portant le même n° y existe déjà, "
                + "il sera écrasé après confirmation.";

            // Toutes les catégories sauf la source.
            var options = new[]
            {
                new OptionTypeMesure(TypeMesure.Frequence,       EnTetesMesureHelper.LibelleType(TypeMesure.Frequence)),
                new OptionTypeMesure(TypeMesure.FreqAvantInterv, EnTetesMesureHelper.LibelleType(TypeMesure.FreqAvantInterv)),
                new OptionTypeMesure(TypeMesure.FreqFinale,      EnTetesMesureHelper.LibelleType(TypeMesure.FreqFinale)),
                new OptionTypeMesure(TypeMesure.Stabilite,       EnTetesMesureHelper.LibelleType(TypeMesure.Stabilite)),
                new OptionTypeMesure(TypeMesure.Interval,        EnTetesMesureHelper.LibelleType(TypeMesure.Interval)),
                new OptionTypeMesure(TypeMesure.TachyContact,    EnTetesMesureHelper.LibelleType(TypeMesure.TachyContact)),
                new OptionTypeMesure(TypeMesure.Stroboscope,     EnTetesMesureHelper.LibelleType(TypeMesure.Stroboscope)),
            };
            var sansSource = options.Where(o => o.Type != categorieSource).ToList();
            CbCategorie.ItemsSource = sansSource;
            if (sansSource.Count > 0) CbCategorie.SelectedIndex = 0;
        }

        private void OnValider(object sender, RoutedEventArgs e)
        {
            if (CbCategorie.SelectedItem is OptionTypeMesure opt)
            {
                CategorieChoisie = opt.Type;
                DialogResult = true;
            }
            else
            {
                System.Windows.MessageBox.Show("Sélectionne une catégorie cible.", "Champ requis",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }

        private void OnAnnuler(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}
