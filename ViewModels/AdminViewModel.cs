using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services.Journal;
using Metrologo.Views;
using System.IO;
using System.Windows;

namespace Metrologo.ViewModels
{
    public partial class AdminViewModel : ObservableObject
    {

        [ObservableProperty] private string _rubidiumActifTexte = EtatApplication.RubidiumActifTexte;
        [ObservableProperty] private string _macroXlaChemin = Preferences.CheminMacroXLA;
        [ObservableProperty] private bool _macroXlaTrouve;
        [ObservableProperty] private string _macroXlaStatut = "";

        // ---------- Format des rapports Excel (commun à tous les postes) ----------

        /// <summary>Réglage enregistré pour tous les postes (relu à chaque ouverture de l'Admin).</summary>
        private FormatRapport _formatEnregistre = Preferences.FormatRapportsPartage;

        /// <summary>Choix en cours dans l'Admin (vrai = .xls) : appliqué seulement par le bouton Appliquer.</summary>
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(RapportsEnXlsx), nameof(ChangementFormatEnAttente), nameof(FormatRapportsStatut))]
        [NotifyCanExecuteChangedFor(nameof(AppliquerFormatRapportsCommand))]
        private bool _rapportsEnXls = Preferences.FormatRapportsPartage == FormatRapport.Xls;

        /// <summary>Inverse de <see cref="RapportsEnXls"/>, pour le second bouton radio.</summary>
        public bool RapportsEnXlsx
        {
            get => !RapportsEnXls;
            set { if (value) RapportsEnXls = false; }
        }

        private FormatRapport FormatChoisi => RapportsEnXls ? FormatRapport.Xls : FormatRapport.Xlsx;

        public bool ChangementFormatEnAttente => FormatChoisi != _formatEnregistre;

        public string FormatRapportsStatut => ChangementFormatEnAttente
            ? $"Format actuel : {Extension(_formatEnregistre)}. Cliquez sur « Appliquer » pour passer en {Extension(FormatChoisi)} : "
              + "Metrologo redémarre sur ce poste, les autres postes passent au nouveau format à leur prochain démarrage."
            : _formatEnregistre == FormatRapport.Xls
                ? "Les rapports sortent en .xls (Excel 97-2003), le format lu par le logiciel externe. Stabilité non concernée."
                : "Les rapports restent en .xlsx. Un ancien rapport .xls repasse en .xlsx à la prochaine mesure sur sa FI.";

        private static string Extension(FormatRapport f) => f == FormatRapport.Xls ? ".xls" : ".xlsx";

        /// <summary>Fourni par la fenêtre principale : vrai pendant une mesure (pas de redémarrage alors).</summary>
        public Func<bool>? MesureEnCours { get; set; }

        /// <summary>À l'ouverture de l'Admin : repart du réglage enregistré (il a pu changer depuis un autre poste).</summary>
        public void RafraichirFormatRapports()
        {
            _formatEnregistre = Preferences.FormatRapportsPartage;
            RapportsEnXls = _formatEnregistre == FormatRapport.Xls;
            OnPropertyChanged(nameof(ChangementFormatEnAttente));
            OnPropertyChanged(nameof(FormatRapportsStatut));
            AppliquerFormatRapportsCommand.NotifyCanExecuteChanged();
        }

        /// <summary>
        /// Enregistre le format pour tous les postes puis redémarre Metrologo sur ce poste (le format
        /// est figé au démarrage). Les autres postes le prennent à leur prochain démarrage.
        /// </summary>
        [RelayCommand(CanExecute = nameof(ChangementFormatEnAttente))]
        private void AppliquerFormatRapports()
        {
            if (MesureEnCours?.Invoke() == true)
            {
                MessageBox.Show("Une mesure est en cours : appliquez le changement de format une fois la mesure terminée.",
                    "Format des rapports", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var format = FormatChoisi;
            var conf = MessageBox.Show(
                $"Passer les rapports Excel en {Extension(format)} pour tous les postes ?\n\n"
              + "• Ce poste : Metrologo va redémarrer maintenant.\n"
              + "• Autres postes : le nouveau format sera pris à leur prochain démarrage de Metrologo.",
                "Format des rapports", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (conf != MessageBoxResult.Yes) return;

            Preferences.FormatRapportsPartage = format;
            Journal.Info(CategorieLog.Administration, "FORMAT_RAPPORTS",
                $"Rapports Excel en sortie de mesure : {Extension(format)} (redémarrage du poste).");

            App.Redemarrer();
        }

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
        private void OuvrirUtilisateurs()
        {
            Journal.Info(CategorieLog.Administration, "OUVERTURE_UTILISATEURS", "Accès à la gestion des utilisateurs.");
            var vm = new GestionUtilisateursViewModel();
            var win = new GestionUtilisateursWindow(vm) { Owner = Application.Current.MainWindow };
            win.ShowDialog();
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
