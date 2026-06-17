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
                    $"{rubi.Designation}",
                    new { id = rubi.Id, gps = rubi.AvecGPS });
            }
        }

        [RelayCommand]
        private void ConfigurerMacroXLA()
        {
            var init = File.Exists(MacroXlaChemin)
                ? Path.GetDirectoryName(MacroXlaChemin) ?? ""
                : @"C:\EXE_SPE\FCT_VBA2016";

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
            catch { /* on ignore, pas grave si l'ouverture echoue */ }
        }

        [RelayCommand]
        private void OuvrirIncertitudes()
        {
            Journal.Info(CategorieLog.Administration, "OUVERTURE_INCERTITUDES", "Consultation du tableau des incertitudes.");
            var win = new DispIncertWindow { Owner = Application.Current.MainWindow };
            win.ShowDialog();
        }

        [RelayCommand]
        private void OuvrirUtilisateurs()
        {
            Journal.Info(CategorieLog.Administration, "OUVERTURE_UTILISATEURS", "Accès à la gestion des utilisateurs.");
            var vm = new GestionUtilisateursViewModel();
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
        /// Ouvre le journal d'AUDIT admin : il trace les actions de configuration (rubidium,
        /// modules d'incertitude, catalogue appareils, utilisateurs…), pas les simples consultations.
        /// </summary>
        [RelayCommand]
        private void OuvrirJournalAdmin()
        {
            var win = new JournalAdminWindow { Owner = Application.Current.MainWindow };
            win.ShowDialog();
        }

        /// <summary>
        /// Lance tout de suite l'archivage du mois précédent, sans attendre le redémarrage de
        /// l'app au 1er du mois. Pratique pour tester, ou pour rattraper un mois passé à la trappe
        /// (PC resté éteint plusieurs jours).
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

        /// <summary>
        /// Déclenche à la main la récupération Besançon (téléchargement FTP du fichier corrigé,
        /// extraction des valeurs, ajout au fichier texte cumulatif), sans attendre l'heure
        /// quotidienne. Rien n'est écrit en base SQL ; le détail part dans le Journal (Système).
        /// </summary>
        [RelayCommand]
        private async System.Threading.Tasks.Task RecupererBesanconMaintenantAsync()
        {
            var conf = MessageBox.Show(
                "Lancer maintenant la récupération du fichier de Besançon sur le FTP "
              + "et l'ajout des valeurs au fichier texte cumulatif ?\n\n"
              + "Nécessite : les identifiants FTP renseignés (fichier besancon.ftp.json).",
                "Récupération Besançon",
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (conf != MessageBoxResult.Yes) return;

            try
            {
                // C'est un lancement manuel : on force, donc on passe outre le garde-fou « deja fait aujourd'hui ».
                var r = await Metrologo.Services.Besancon.BesanconScheduler.ExecuterAsync(forcer: true);

                if (!r.Succes)
                {
                    MessageBox.Show(
                        "Récupération Besançon NON aboutie.\n\n" + (r.Erreur ?? "Cause inconnue — voir le Journal (Système)."),
                        "Besançon", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                MessageBox.Show(
                    "Récupération Besançon terminée ✔\n\n"
                  + $"• Valeurs lues dans le fichier FTP : {r.ValeursLues}\n"
                  + $"• Nouvelles valeurs ajoutées : {r.Nouvelles}\n"
                  + $"• Total dans le fichier cumulatif : {r.TotalJournalieres}\n\n"
                  + $"Fichier cumulatif :\n{(r.EnregistrementOk ? r.Destination : "⚠ ÉCHEC — voir le Journal")}\n\n"
                  + $"Fichier brut (copie datée) :\n{r.CheminBrut ?? "⚠ NON ÉCRIT — voir le Journal (partage injoignable ?)"}",
                    "Besançon", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Récupération Besançon échouée : {ex.Message}",
                    "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        /// <summary>Ouvre les reglages de la recuperation automatique Besancon (activation, heure, FTP).</summary>
        [RelayCommand]
        private void OuvrirParametresBesancon()
        {
            Journal.Info(CategorieLog.Administration, "OUVERTURE_PARAMETRES_BESANCON",
                "Accès aux paramètres de la récupération automatique Besançon.");
            var win = new BesanconParametresWindow { Owner = Application.Current.MainWindow };
            win.ShowDialog();
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
            // On est dans AdminViewModel : forcement un admin derriere (l'acces a AdminView
            // n'est propose qu'aux comptes Administrateur).
            var vm = new GestionAppareilsViewModel(utilisateur, estAdmin: true);
            var win = new GestionAppareilsWindow(vm) { Owner = Application.Current.MainWindow };
            win.ShowDialog();
        }

        /// <summary>
        /// Ouvre l'editeur des adresses GPIB des appareils legacy (EIP / Racal / Stanford).
        /// Reserve a l'admin : on y attribue une adresse distincte a chaque appareil pour pouvoir
        /// les brancher en meme temps. Tout est enregistre sur le reseau (appareils-legacy.json).
        /// </summary>
        [RelayCommand]
        private void OuvrirAdressesLegacy()
        {
            Journal.Info(CategorieLog.Administration, "OUVERTURE_ADRESSES_LEGACY",
                "Accès à la configuration des adresses GPIB des appareils fixes.");

            var vm = new GestionAdressesLegacyViewModel();
            var win = new GestionAdressesLegacyWindow(vm) { Owner = Application.Current.MainWindow };
            win.ShowDialog();
        }

        /// <summary>
        /// Ouvre la configuration des chemins de stockage. L'admin peut y remplacer les
        /// emplacements par defaut (locaux) par un partage reseau commun a tous les postes du
        /// site (modules d'incertitude, presets, archives logs, etc.). La base SQL Server, elle,
        /// reste centralisee et ne se configure pas ici (voir CheminsMetrologo + paths.config.json).
        /// </summary>
        [RelayCommand]
        private void OuvrirCheminsStockage()
        {
            Journal.Info(CategorieLog.Administration, "OUVERTURE_CHEMINS_STOCKAGE",
                "Accès à la configuration des chemins de stockage.");

            var vm = new CheminsStockageViewModel();
            var win = new CheminsStockageWindow(vm) { Owner = Application.Current.MainWindow };
            win.ShowDialog();
        }
    }
}
