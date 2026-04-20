using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Services.Journal;
using Metrologo.Views;
using System.Windows;

namespace Metrologo.ViewModels
{
    public partial class AdminViewModel : ObservableObject
    {
        [ObservableProperty]
        private string _messageAdmin =
            "Bienvenue dans l'espace administrateur. Gérez les utilisateurs, "
            + "les coefficients d'incertitude, la configuration système et le journal.";

        [RelayCommand]
        private void OuvrirIncertitudes()
        {
            Journal.Info(CategorieLog.Administration, "OUVERTURE_INCERTITUDES", "Consultation du tableau des incertitudes.");
            var win = new DispIncertWindow { Owner = Application.Current.MainWindow };
            win.ShowDialog();
        }

        [RelayCommand]
        private void OuvrirParamsIncert()
        {
            Journal.Info(CategorieLog.Administration, "OUVERTURE_PARAMS_INCERT", "Ouverture du dialogue résolution / incert. supp.");
            var vm = new ParamsIncertViewModel(0.01, 0.0);
            var win = new ParamsIncertWindow(vm) { Owner = Application.Current.MainWindow };
            if (win.ShowDialog() == true)
            {
                Journal.Info(CategorieLog.Administration, "PARAMS_INCERT_MAJ",
                    $"Résolution {vm.Resolution} · Incert. supp. {vm.IncertSupp}",
                    new { vm.Resolution, vm.IncertSupp });
            }
        }

        [RelayCommand]
        private void OuvrirUtilisateurs()
        {
            Journal.Info(CategorieLog.Administration, "OUVERTURE_UTILISATEURS", "Accès à la gestion des utilisateurs.");
            MessageBox.Show("Gestion des utilisateurs — à implémenter.",
                "En développement", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        [RelayCommand]
        private void OuvrirSysteme()
        {
            Journal.Info(CategorieLog.Administration, "OUVERTURE_SYSTEME", "Accès à la configuration système.");
            MessageBox.Show("Configuration système — à implémenter.",
                "En développement", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        [RelayCommand]
        private void OuvrirJournal()
        {
            Journal.Info(CategorieLog.Administration, "OUVERTURE_JOURNAL", "Consultation du journal.");
            var win = new JournalViewerWindow { Owner = Application.Current.MainWindow };
            win.ShowDialog();
        }
    }
}
