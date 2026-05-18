using System.ComponentModel;
using System.Windows;
using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class GestionModulesIncertitudeWindow : FluentWindow
    {
        public GestionModulesIncertitudeWindow()
        {
            InitializeComponent();
            var vm = new GestionModulesIncertitudeViewModel();
            DataContext = vm;
            vm.PropertyChanged += OnVmPropertyChanged;
            AppliquerVisibiliteColTemps(vm);
        }

        // Masque la colonne « Temps (s) » du DataGrid quand le module sélectionné a
        // UtiliseTempsDeMesure = false (cas tachymètre/stroboscope). DataGridColumn
        // n'hérite pas du DataContext donc on pilote la Visibility en code-behind.
        private void OnVmPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(GestionModulesIncertitudeViewModel.ModuleSelectionne)
                && DataContext is GestionModulesIncertitudeViewModel vm)
            {
                AppliquerVisibiliteColTemps(vm);
            }
        }

        private void AppliquerVisibiliteColTemps(GestionModulesIncertitudeViewModel vm)
        {
            bool afficher = vm.ModuleSelectionne?.UtiliseTempsDeMesure ?? true;
            ColTemps.Visibility = afficher ? Visibility.Visible : Visibility.Collapsed;
        }

        private void OnAideSaisie(object sender, RoutedEventArgs e)
        {
            var w = new AideSaisieIncertitudeWindow { Owner = this };
            w.ShowDialog();
        }
    }
}
