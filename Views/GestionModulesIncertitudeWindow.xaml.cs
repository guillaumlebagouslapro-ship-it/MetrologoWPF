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

        // Colonne « Temps (s) » masquée si UtiliseTempsDeMesure = false (tachymètre/stroboscope).
        // Une DataGridColumn n'hérite pas du DataContext : Visibility réglée ici, en code-behind.
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
