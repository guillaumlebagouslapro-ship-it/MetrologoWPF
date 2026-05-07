using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services;
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
            var vm = new GestionUtilisateursViewModel(new UtilisateursService());
            var win = new GestionUtilisateursWindow(vm) { Owner = Application.Current.MainWindow };
            win.ShowDialog();
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

        /// <summary>
        /// Force l'archivage du mois précédent immédiatement (sans attendre que l'app
        /// redémarre le 1er du mois). Utile pour test ou pour rattraper un mois qui
        /// aurait été manqué (PC éteint pendant plusieurs jours).
        /// </summary>
        [RelayCommand]
        private async System.Threading.Tasks.Task ArchiverMaintenantAsync()
        {
            var maintenant = DateTime.Now;
            var moisPrec = new DateTime(maintenant.Year, maintenant.Month, 1).AddMonths(-1);

            var conf = MessageBox.Show(
                $"Archiver les logs de {moisPrec:MMMM yyyy} maintenant ?\n\n" +
                "Les logs de ce mois seront exportés dans des fichiers JSON " +
                $"({Metrologo.Services.Journal.ArchivesLogsService.DossierArchivesRacine}) " +
                "puis supprimés de la base SQL Server.",
                "Archiver les logs",
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (conf != MessageBoxResult.Yes) return;

            try
            {
                int n = await Metrologo.Services.Journal.ArchivesLogsService
                    .ArchiverMoisAsync(moisPrec, force: true);
                MessageBox.Show(
                    $"Archivage terminé : {n} entrée(s) exportée(s).\n\n" +
                    $"Dossier : {Metrologo.Services.Journal.ArchivesLogsService.DossierArchivesRacine}\\{moisPrec:yyyy-MM}",
                    "Archivage OK", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Archivage échoué : {ex.Message}",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        [RelayCommand]
        private void OuvrirGestionModules()
        {
            Journal.Info(CategorieLog.Administration, "OUVERTURE_GESTION_MODULES",
                "Accès à la gestion des modules d'incertitude.");
            var win = new GestionModulesIncertitudeWindow { Owner = Application.Current.MainWindow };
            win.ShowDialog();
        }

        [RelayCommand]
        private void OuvrirGestionAppareils()
        {
            Journal.Info(CategorieLog.Administration, "OUVERTURE_GESTION_APPAREILS",
                "Accès à la gestion du catalogue d'appareils.");

            string utilisateur = Journal.Utilisateur ?? "admin";
            // Cette commande est dans AdminViewModel, donc forcément lancée par un admin
            // (la nav vers AdminView n'est exposée qu'aux comptes Administrateur).
            var vm = new GestionAppareilsViewModel(utilisateur, estAdmin: true);
            var win = new GestionAppareilsWindow(vm) { Owner = Application.Current.MainWindow };
            win.ShowDialog();
        }
    }
}
