using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services.Journal;
using Metrologo.Views;
using System;
using System.IO;
using System.Windows;

namespace Metrologo.ViewModels
{
    public partial class AdminViewModel : ObservableObject
    {
        [ObservableProperty]
        private string _messageAdmin =
            "Bienvenue dans l'espace administrateur. Gérez les utilisateurs, "
            + "les coefficients d'incertitude, la configuration système et le journal.";

        [ObservableProperty] private string _rubidiumActifTexte = EtatApplication.RubidiumActifTexte;
        [ObservableProperty] private string _macroXlaChemin = Preferences.CheminMacroXLA;
        [ObservableProperty] private bool _macroXlaTrouve;
        [ObservableProperty] private string _macroXlaStatut = "";

        public AdminViewModel()
        {
            EtatApplication.RubidiumActifChange += OnRubidiumChange;
            MajStatutMacro();
        }

        private void OnRubidiumChange(object? sender, EventArgs e)
        {
            RubidiumActifTexte = EtatApplication.RubidiumActifTexte;
        }

        private void MajStatutMacro()
        {
            MacroXlaChemin = Preferences.CheminMacroXLA;
            MacroXlaTrouve = File.Exists(MacroXlaChemin);
            MacroXlaStatut = MacroXlaTrouve
                ? "Macro trouvée — les formules d'incertitude seront calculées"
                : "Macro introuvable — les formules afficheront #NOM? à l'ouverture d'Excel";
        }

        [RelayCommand]
        private void ChoisirRubidium()
        {
            var win = new ChoixRubidiumWindow { Owner = Application.Current.MainWindow };
            if (win.ShowDialog() == true && win.ViewModel.Resultat != null)
            {
                var rubi = win.ViewModel.Resultat;
                EtatApplication.RubidiumActif = rubi;

                Journal.Info(CategorieLog.Rubidium, "SELECTION_RUBIDIUM",
                    $"{rubi.Designation} — {(rubi.AvecGPS ? "GPS" : "Allouis")}",
                    new { id = rubi.Id, gps = rubi.AvecGPS });
            }
        }

        [RelayCommand]
        private void ConfigurerMacroXLA()
        {
            var init = File.Exists(MacroXlaChemin)
                ? Path.GetDirectoryName(MacroXlaChemin) ?? ""
                : @"C:\Exe_Spe\Fct_VBA";

            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Macro Excel (*.xla;*.xlam)|*.xla;*.xlam|Tous les fichiers|*.*",
                Title = "Sélectionner Metrologo.xla",
                InitialDirectory = Directory.Exists(init) ? init : ""
            };

            if (dlg.ShowDialog() == true)
            {
                Preferences.CheminMacroXLA = dlg.FileName;
                MajStatutMacro();
                Journal.Info(CategorieLog.Administration, "MACRO_XLA_CONFIG",
                    $"Chemin Metrologo.xla : {dlg.FileName}");
            }
        }

        [RelayCommand]
        private void OuvrirDossierMacroXLA()
        {
            try
            {
                var dir = Path.GetDirectoryName(MacroXlaChemin);
                if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = dir,
                        UseShellExecute = true
                    });
                }
            }
            catch { /* silencieux */ }
        }

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
