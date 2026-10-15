using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Services.Besancon;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Ieee;
using Metrologo.Services.Journal;
using Metrologo.Views;
using System.Windows;
using System.Windows.Media;

namespace Metrologo.ViewModels
{
    public partial class AccueilViewModel : ObservableObject
    {
        // Driver IEEE bas niveau : VISA réel (NI-VISA, matériel branché).
        // Sans matériel, basculer sur new SimulationIeeeDriver().
        private readonly IIeeeDriver _ieeeDriver = new VisaIeeeDriver(gpibBoard: 0);
        // Type concret (et pas IExcelService) pour pouvoir lire FallbackTimestampUtilise.
        private readonly ExcelService _excelService = new ExcelService();
        private readonly MesureOrchestrator _orchestrator;

        private CancellationTokenSource? _cts;

        /// <summary>Token d'abandon force : si on clique STOP alors que la tache est coincee sur un
        /// COM Excel mort (RPC ~60 s), il rend la main a l'UI sans attendre la fin de l'orchestration.</summary>
        private CancellationTokenSource? _abandonCts;

        /// <summary>Vrai pendant toute une série de stabs répétées : bloque Lancer/Relancer entre deux
        /// répétitions (MesureEnCours repasse brièvement à false entre la fin d'une stab et la suivante).</summary>
        private bool _serieStabEnCours;

        /// <summary>Position dans la série de stabs en cours (1/1 hors série) : sert à l'heure de fin
        /// affichée dans le bandeau rouge (<see cref="SuiviFinStab"/>).</summary>
        private int _repetitionStab = 1;
        private int _nbRepetitionsStab = 1;

        /// <summary>Nombre de stabs enchaînées au dernier lancement : « Relancer » rejoue la même série.</summary>
        private int _derniereNbRepetitionsStab = 1;

        /// <summary>Vrai tant qu'une mesure tourne, s'arrête, ou qu'une série de stabs est en cours
        /// (y compris entre deux répétitions). Sert à ne jamais déconnecter en pleine acquisition.</summary>
        public bool ActiviteMesureEnCours => MesureEnCours || ArretEnCours || _serieStabEnCours;

        [ObservableProperty] private bool _estSurBaie = true;
        [ObservableProperty] private string _informationsGenerales = "Prêt. En attente d'exécution...";
        [ObservableProperty] private string _rubidiumActifTexte = EtatApplication.RubidiumActifTexte;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(EstConfigure))]
        [NotifyPropertyChangedFor(nameof(LibelleAppareilConfigure))]
        private Mesure _mesureConfig = new Mesure();

        /// <summary>Vrai des qu'une config est validee (N° FI renseigne). Pilote le bandeau « Fiche en cours » de l'accueil.</summary>
        public bool EstConfigure => !string.IsNullOrWhiteSpace(MesureConfig?.NumFI);

        /// <summary>Nom court de l'appareil de la config (vide si aucun), affiche dans le bandeau « Fiche en cours ».</summary>
        public string LibelleAppareilConfigure
        {
            get
            {
                if (string.IsNullOrEmpty(MesureConfig?.IdModeleCatalogue)) return string.Empty;
                var modele = CatalogueAppareilsService.Instance.Modeles
                    .FirstOrDefault(m => m.Id == MesureConfig.IdModeleCatalogue);
                return modele?.Nom ?? string.Empty;
            }
        }

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(RelancerMesureCommand))]
        [NotifyCanExecuteChangedFor(nameof(StopperMesureCommand))]
        private bool _mesureEnCours;

        /// <summary>Vrai entre le clic sur « Arreter » et la fin reelle de la mesure (token annule + Device Clear).
        /// Donne un retour immediat a l'utilisateur, le temps que le compteur libere son :FETCh? en cours.</summary>
        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(StopperMesureCommand))]
        private bool _arretEnCours;

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(RelancerMesureCommand))]
        private bool _derniereMesureDisponible;

        /// <summary>Nombre d'appareils detectes sur le bus GPIB, reactualise apres chaque scan.</summary>
        [ObservableProperty] private int _nbAppareilsDetectes;

        /// <summary>Resume du scan ("3 detecte(s), 2 reconnu(s)") affiche sur la carte Diagnostic.</summary>
        [ObservableProperty] private string _resumeScanGpib = "Scan en cours...";

        /// <summary>Vrai tant que le scan initial tourne (masque le texte detaille).</summary>
        [ObservableProperty] private bool _scanInitialEnCours = true;

        private double? _derniereFNominale;

        /// <summary>Rapports pour lesquels la question « écraser / ajouter » a déjà été posée (ou
        /// qui ont été créés) pendant cette session : les mesures suivantes de la série (feuilles
        /// 2, 3…) s'y ajoutent sans redemander.</summary>
        private readonly HashSet<string> _rapportsConfirmes = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>FI déjà signalées « restée en local » cette session (message une seule fois).</summary>
        private readonly HashSet<string> _fiAvertiesLocal = new(StringComparer.OrdinalIgnoreCase);

        // ---- Suivi Besancon (panneau d'etat + voyant sur l'ecran principal) ----

        /// <summary>Affiche le panneau de suivi Besancon (vrai des qu'un rubidium actif est defini).</summary>
        [ObservableProperty] private bool _besanconVisible;

        /// <summary>Titre du panneau (A jour, Retard, Critique...).</summary>
        [ObservableProperty] private string _besanconTitre = "Suivi Besançon";

        /// <summary>Message detaille sous le titre (anciennete des donnees, cause d'un probleme…).</summary>
        [ObservableProperty] private string _besanconDetail = "Chargement du suivi…";

        /// <summary>Couleur du voyant : vert / orange / rouge, ou gris quand l'etat est indetermine.</summary>
        [ObservableProperty] private Brush _besanconVoyant = new SolidColorBrush(Color.FromRgb(0x9C, 0xA3, 0xAF));

        /// <summary>Rapport texte indente (valeurs journalieres + moyennes hebdo) affiche dans le panneau.</summary>
        [ObservableProperty] private string _besanconRapport = string.Empty;

        public AccueilViewModel()
        {
            _orchestrator = new MesureOrchestrator(_ieeeDriver, _excelService);

            // Suit le changement de rubidium actif par l'administrateur.
            EtatApplication.RubidiumActifChange += (_, _) =>
            {
                RubidiumActifTexte = EtatApplication.RubidiumActifTexte;
                _ = RafraichirBesanconAsync();   // rubidium change => on reevalue le suivi
            };

            // Idem si un scan est relance depuis Diagnostic GPIB (ajout d'un appareil).
            EtatApplication.AppareilsDetectesChange += (_, _) => RafraichirResumeScan();

            // Voyant et rapport rafraîchis après chaque passage de la tâche Besançon.
            BesanconScheduler.StatutChange += (_, _) => _ = RafraichirBesanconAsync();

            // Scan GPIB initial en tâche de fond (ne doit pas bloquer l'ouverture de la fenêtre).
            _ = Task.Run(ScannerInitialAsync);

            // État Besançon de départ.
            _ = RafraichirBesanconAsync();
        }

        /// <summary>Reevalue le suivi Besancon (voyant + rapport) depuis la base partagee. Appele au
        /// demarrage, au changement de rubidium et apres chaque passage de la tache. Repasse sur le thread UI.</summary>
        [RelayCommand]
        private async Task RafraichirBesanconAsync()
        {
            var rub = EtatApplication.RubidiumActif;
            BesanconStatut st;
            try { st = await BesanconSuiviService.EvaluerAsync(rub, DateTime.Today); }
            catch (Exception ex)
            {
                st = new BesanconStatut
                {
                    Niveau = NiveauSuivi.Inconnu,
                    Titre = "Suivi Besançon — Indisponible",
                    Detail = ex.Message,
                    RapportTxt = string.Empty
                };
            }

            void Appliquer()
            {
                BesanconVisible = rub != null;
                BesanconTitre = st.Titre;
                BesanconDetail = st.Detail;
                BesanconRapport = st.RapportTxt;
                BesanconVoyant = new SolidColorBrush(st.Niveau switch
                {
                    NiveauSuivi.Vert   => Color.FromRgb(0x16, 0xA3, 0x4A),
                    NiveauSuivi.Orange => Color.FromRgb(0xF5, 0x9E, 0x0B),
                    NiveauSuivi.Rouge  => Color.FromRgb(0xDC, 0x26, 0x26),
                    _                  => Color.FromRgb(0x9C, 0xA3, 0xAF),
                });
            }

            var disp = Application.Current?.Dispatcher;
            if (disp != null && !disp.CheckAccess()) disp.Invoke(Appliquer);
            else Appliquer();
        }

        /// <summary>Relance manuelle du scan GPIB (bouton « Rescanner »), pratique quand on branche un appareil en cours de session.</summary>
        [RelayCommand]
        private Task RescannerAsync()
        {
            Log("🔄 Relance du scan GPIB...");
            return ScannerInitialAsync();
        }

        /// <summary>Scan GPIB automatique au demarrage : il alimente NbAppareilsDetectes et le resume dans InformationsGenerales.</summary>
        private async Task ScannerInitialAsync()
        {
            try
            {
                ScanInitialEnCours = true;
                ResumeScanGpib = "Scan en cours...";

                Journal.Info(CategorieLog.Systeme, "SCAN_AUTO_DEBUT",
                    "Scan GPIB automatique au démarrage de Metrologo.");

                // Timeout de 3 s : certains appareils (53131A surtout) répondent lentement au *IDN?
                // après une mise sous tension ou une perturbation du bus.
                var resultats = await ScannerGpib.ScannerAsync(gpibBoard: 0, timeoutMs: 3000);

                // État global des appareils détectés (lu ensuite par Configuration, etc.).
                var appareilsDetectes = resultats
                    .Where(r => r.Repond)
                    .Select(r =>
                    {
                        var modeleReconnu = CatalogueAppareilsService.Instance
                            .TrouverParIdn(r.Fabricant, r.Modele);
                        return new AppareilDetecte
                        {
                            Board = r.Board,
                            Adresse = r.Adresse,
                            Ressource = r.Ressource,
                            TypeBus = r.TypeBus,
                            Hote = r.Hote,
                            IdnBrut = r.ReponseIdn,
                            Fabricant = r.Fabricant,
                            Modele = r.Modele,
                            NumeroSerie = r.NumeroSerie,
                            Firmware = r.Firmware,
                            ModeleReconnu = modeleReconnu,
                            ConflitAdressePossible = r.ConflitAdressePossible
                        };
                    })
                    .ToList();

                Application.Current.Dispatcher.Invoke(() =>
                {
                    EtatApplication.AppareilsDetectes.Clear();
                    foreach (var app in appareilsDetectes)
                        EtatApplication.AppareilsDetectes.Add(app);
                    EtatApplication.NotifierAppareilsDetectesChange();

                    ScanInitialEnCours = false;
                    RafraichirResumeScan();
                    AfficherAppareilsDansInformations();
                });

                Journal.Info(CategorieLog.Systeme, "SCAN_AUTO_FIN",
                    $"Scan auto terminé : {appareilsDetectes.Count} appareil(s) détecté(s).",
                    new { NbDetectes = appareilsDetectes.Count });
            }
            catch (Exception ex)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    ScanInitialEnCours = false;
                    ResumeScanGpib = "Scan échoué";
                    Log($"⚠ Scan GPIB initial impossible : {ex.Message}");
                });

                Journal.Erreur(CategorieLog.Systeme, "SCAN_AUTO_ERREUR",
                    $"Scan GPIB initial échoué : {ex.Message}",
                    new { ex.GetType().Name });
            }
        }

        private void RafraichirResumeScan()
        {
            int n = EtatApplication.AppareilsDetectes.Count;
            int nReconnus = EtatApplication.AppareilsDetectes.Count(a => a.EstReconnu);

            NbAppareilsDetectes = n;
            ResumeScanGpib = n == 0
                ? "Aucun appareil"
                : nReconnus == n
                    ? $"{n} détecté · {n} reconnu{(n > 1 ? "s" : "")}"
                    : $"{n} détecté{(n > 1 ? "s" : "")} · {nReconnus} reconnu{(nReconnus > 1 ? "s" : "")}";
        }

        private void AfficherAppareilsDansInformations()
        {
            var list = EtatApplication.AppareilsDetectes.ToList();
            if (list.Count == 0)
            {
                Log("ℹ Aucun appareil GPIB détecté sur le bus.");
                return;
            }

            Log($"🔌 Scan GPIB initial : {list.Count} appareil(s) détecté(s) :");
            foreach (var app in list)
            {
                string statut = app.ModeleReconnu != null
                    ? $"catalogue : {app.ModeleReconnu.Nom}"
                    : "non enregistré";
                Log($"   • {app.AdresseCourte} — {app.Fabricant} {app.Modele} ({statut})");
            }

            // Une réponse *IDN? incohérente trahit souvent un conflit d'adresse (2 appareils sur la même).
            var conflits = list.Where(a => a.ConflitAdressePossible).ToList();
            if (conflits.Count > 0)
            {
                foreach (var c in conflits)
                {
                    Log($"⚠ {c.AdresseCourte} : réponse d'identification incohérente — "
                      + "deux appareils sont peut-être réglés sur la même adresse GPIB. "
                      + "Vérifie que chaque instrument a une adresse unique.");
                    Journal.Warn(CategorieLog.Systeme, "GPIB_CONFLIT_ADRESSE",
                        $"Réponse *IDN? incohérente à {c.AdresseCourte} (IDN brut : « {c.IdnBrut} ») "
                      + "— possible conflit d'adresse (2 appareils sur la même adresse).");
                }
            }
        }

        // -------- Commandes --------

        /// <summary>Ouvre dans l'Explorateur le dossier des fichiers Excel generes (cree au besoin).</summary>
        [RelayCommand]
        private void OuvrirDossierMesures()
        {
            // Chemin réseau configuré (Admin > Chemins d'accès > Mesures), sinon Bureau\Metrologo
            // si le réseau est vide ou injoignable.
            string dossierLocal = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "Metrologo");
            string dossierReseau = Services.CheminsMetrologo.MesuresLocal;

            string dossierCible = dossierLocal;
            if (!string.IsNullOrWhiteSpace(dossierReseau))
            {
                try
                {
                    if (System.IO.Directory.Exists(dossierReseau)) dossierCible = dossierReseau;
                }
                catch { /* reseau injoignable : on reste en local */ }
            }

            try
            {
                System.IO.Directory.CreateDirectory(dossierCible);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = dossierCible,
                    UseShellExecute = true
                });
                Journal.Info(CategorieLog.Systeme, "OUVERTURE_DOSSIER_MESURES",
                    $"Ouverture du dossier des mesures : {dossierCible}");
            }
            catch (Exception ex)
            {
                Log($"⚠ Impossible d'ouvrir le dossier : {ex.Message}");
                Journal.Erreur(CategorieLog.Systeme, "OUVERTURE_DOSSIER_MESURES_ERR",
                    $"Échec : {ex.Message}", new { ex.GetType().Name });
            }
        }

        [RelayCommand]
        private void OuvrirDiagnosticGpib()
        {
            Journal.Info(CategorieLog.Systeme, "OUVERTURE_DIAGNOSTIC_GPIB",
                "Accès au diagnostic du bus GPIB depuis l'accueil.");
            var win = new DiagnosticGpibWindow { Owner = Application.Current.MainWindow };
            win.ShowDialog();
        }

        [RelayCommand]
        private async Task OuvrirConfigurationAsync()
        {
            if (MesureConfig == null) MesureConfig = new Mesure();

            var vm = new ConfigurationViewModel { MesureConfig = MesureConfig };
            vm.EstSurBaie = EstSurBaie;

            var win = new ConfigurationWindow(vm) { Owner = Application.Current.MainWindow };
            if (win.ShowDialog() == true)
            {
                MesureConfig = vm.MesureConfig;
                // Même référence réaffectée (le vm a reçu notre instance) : CommunityToolkit ne notifie
                // pas, d'où la notification forcée pour le bandeau « Fiche en cours ».
                OnPropertyChanged(nameof(MesureConfig));
                OnPropertyChanged(nameof(EstConfigure));
                OnPropertyChanged(nameof(LibelleAppareilConfigure));
                string nomAppareil = vm.AppareilSelectionne?.Detecte?.Libelle ?? "(aucun)";
                Log($"⚙ Configuration : FI {MesureConfig.NumFI} · {MesureConfig.TypeMesure} · "
                  + $"{nomAppareil} · {MesureConfig.NbMesures} mesures · "
                  + $"Mode {MesureConfig.ModeMesure}");

                Journal.Info(CategorieLog.Configuration, "CONFIG_VALIDEE",
                    $"FI {MesureConfig.NumFI} · {MesureConfig.TypeMesure} · {nomAppareil}",
                    new
                    {
                        numFI = MesureConfig.NumFI,
                        type = MesureConfig.TypeMesure.ToString(),
                        appareil = nomAppareil,
                        idModele = MesureConfig.IdModeleCatalogue,
                        nbMesures = MesureConfig.NbMesures,
                        mode = MesureConfig.ModeMesure.ToString(),
                        source = MesureConfig.SourceMesure.ToString(),
                        gateIndex = MesureConfig.GateIndex,
                        fNominale = MesureConfig.FNominale
                    });

                // Enchainement automatique : une fois la config validee, on lance directement
                // la mesure (ouvrirConfigAvant=false pour ne pas rouvrir la fenetre de config en boucle).
                await ExecuterMesureInterneAsync(ouvrirConfigAvant: false);
            }
        }

        /// <summary>Bouton « Lancer la mesure » : passe toujours par la fenêtre de config, qui rappelle
        /// ExecuterMesureInterneAsync(false) une fois validée. Sans reconfigurer : bouton « Relancer ».</summary>
        [RelayCommand]
        private Task ExecuterMesureAsync() => ExecuterMesureInterneAsync(ouvrirConfigAvant: true);

        private async Task ExecuterMesureInterneAsync(bool ouvrirConfigAvant)
        {
            if (MesureEnCours || _serieStabEnCours) return;

            // 1) Rubidium obligatoire (défini uniquement dans le menu Administration).
            var rubi = EtatApplication.RubidiumActif;
            if (rubi == null)
            {
                Log("✖ Mesure impossible : aucun rubidium actif.");
                Journal.Warn(CategorieLog.Mesure, "MESURE_BLOQUEE",
                    "Tentative de mesure sans rubidium actif.");
                MessageBox.Show(
                    "Aucun rubidium n'est défini comme actif.\n\n"
                    + "Un administrateur doit en définir un via le menu Administration.",
                    "Rubidium requis", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 2) « Lancer la mesure » ouvre la config même si une config valide existe déjà ; à la validation,
            //    OuvrirConfigurationAsync rappelle ExecuterMesureInterneAsync(false) : d'où le return ici.
            if (ouvrirConfigAvant || string.IsNullOrWhiteSpace(MesureConfig?.NumFI))
            {
                Log("ℹ Ouverture de la configuration avant lancement.");
                await OuvrirConfigurationAsync();
                return;
            }

            // 2bis) Module d'incertitude obligatoire : sans lui, les coefficients A/B (C/D en tachy) restent
            // à des valeurs par défaut codées en dur et faussent l'incertitude du rapport.
            if (string.IsNullOrWhiteSpace(MesureConfig.NumModuleIncertitude))
            {
                Log("✖ Mesure impossible : aucun module d'incertitude sélectionné.");
                Journal.Warn(CategorieLog.Mesure, "MESURE_BLOQUEE_MODULE",
                    "Tentative de mesure sans module d'incertitude sélectionné.");
                MessageBox.Show(
                    "Aucun module d'incertitude n'est sélectionné dans la configuration.\n\n"
                    + "Choisis-en un dans le menu déroulant « MODULE D'INCERTITUDE » avant de relancer.\n\n"
                    + "Si la liste est vide, va dans Admin → Modules d'incertitude pour en créer un "
                    + "pour ce type de mesure.",
                    "Module requis", MessageBoxButton.OK, MessageBoxImage.Warning);
                await OuvrirConfigurationAsync();
                return;
            }

            // 2ter) En tachymétrie, le module Fréquence auxiliaire (ZNCoeffA/B en Hz) est aussi obligatoire :
            // il caractérise le fréquencemètre support, sans lui le rapport est incomplet côté Hz.
            if (EnTetesMesureHelper.EstTachymetre(MesureConfig.TypeMesure)
                && string.IsNullOrWhiteSpace(MesureConfig.NumModuleIncertitudeFreq))
            {
                Log("✖ Mesure tachy/strobo impossible : module Fréquence auxiliaire non sélectionné.");
                Journal.Warn(CategorieLog.Mesure, "MESURE_BLOQUEE_MODULE_FREQ",
                    "Tentative de mesure tachy/strobo sans module Fréquence auxiliaire.");
                MessageBox.Show(
                    "Pour une mesure tachymétrique ou stroboscope, il faut sélectionner DEUX modules :\n"
                    + "  • Module d'incertitude (tachy/strobo) — pour les Coeff C/D côté RPM\n"
                    + "  • Module Fréquence (Hz) — pour les Coeff A/B côté Hz (fréquencemètre)\n\n"
                    + "Le second est manquant. Va dans la fenêtre Configuration et choisis "
                    + "le module Fréquence correspondant au compteur utilisé.",
                    "Module Fréquence requis", MessageBoxButton.OK, MessageBoxImage.Warning);
                await OuvrirConfigurationAsync();
                return;
            }

            // 3) Gate : déjà choisie en Configuration (MesureConfig.GateIndex). En Stabilité, fenêtre dédiée
            //    pour choisir les gates à balayer (cases + presets) et le nombre de répétitions.
            int nbRepetitions = 1;
            if (MesureConfig.TypeMesure == TypeMesure.Stabilite)
            {
                var gateWin = new SelectionGateWindow(MesureConfig) { Owner = Application.Current.MainWindow };
                if (gateWin.ShowDialog() != true) { Log("✖ Mesure annulée (sélection gates)."); return; }
                MesureConfig.GateIndices = gateWin.ViewModel.IndicesGatesResultats;
                nbRepetitions = gateWin.ViewModel.NbRepetitions;
                _derniereNbRepetitionsStab = Math.Max(1, nbRepetitions);
            }
            else if (MesureConfig.GateIndices.Count > 1)
            {
                // Filet de sécurité : après une Stab, ses gates multiples traînent en mémoire ; on réduit à
                // la 1ère, sinon la mesure Fréquence boucle sur toutes les gates de la Stab précédente.
                MesureConfig.GateIndex = MesureConfig.GateIndices[0]; // le setter ramene la liste a 1 element
            }
            Log($"⏱ Gates à balayer : {string.Join(", ", MesureConfig.GateIndices)}");

            // 3bis) Anti-collision de FI : si un rapport du MÊME nom existe déjà (N° repris d'une session
            //    précédente), on demande à l'utilisateur (sinon l'ancien freq.xlsx serait réutilisé en silence,
            //    avec une 2e instance Excel à fenêtre grise si il était resté ouvert). Même choix que
            //    « Relancer ». Posée UNE fois par rapport et par session (feuilles 2, 3… sans redemander).
            //    Stabilité exclue : elle versionne déjà (stab1, stab2…) et n'écrase jamais.
            if (MesureConfig.TypeMesure != TypeMesure.Stabilite)
            {
                // Hors thread UI : le choix du dossier de la FI peut tester le réseau. Une FI passée
                // en local revient d'abord sur le réseau s'il répond de nouveau.
                var config = MesureConfig;
                await CheminsMetrologo.ReevaluerDossierFIAsync(config.NumFI ?? string.Empty);
                string cheminAttendu = await Task.Run(() => _excelService.CalculerCheminFichierAttendu(config));
                // Le rapport peut être sorti en .xls (réglage Admin) : il compte aussi.
                string cheminXlsAttendu = System.IO.Path.ChangeExtension(cheminAttendu, ".xls");
                string? cheminExistant = System.IO.File.Exists(cheminAttendu) ? cheminAttendu
                    : System.IO.File.Exists(cheminXlsAttendu) ? cheminXlsAttendu : null;
                if (cheminExistant != null && !_rapportsConfirmes.Contains(cheminAttendu))
                {
                    var choix = MessageBox.Show(
                        $"Un rapport existe déjà pour la FI {MesureConfig.NumFI} :\n{cheminExistant}\n\n"
                      + "Voulez-vous écraser les mesures précédentes ?\n\n"
                      + "• Oui : repart d'un fichier vierge (les anciennes mesures sont perdues).\n"
                      + "• Non : conserve l'historique et ajoute la nouvelle mesure à la suite.\n"
                      + "• Annuler : ne lance rien.",
                        "Rapport déjà existant pour cette FI",
                        MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);

                    if (choix == MessageBoxResult.Cancel)
                    {
                        Log("✖ Mesure annulée : un rapport existe déjà pour cette FI.");
                        return;
                    }

                    if (choix == MessageBoxResult.Yes && !await EcraserRapportExistantAsync(cheminAttendu))
                        return;   // suppression impossible (fichier verrouillé) : message déjà affiché
                }
                _rapportsConfirmes.Add(cheminAttendu);
            }

            // 4) Fréquence nominale déjà saisie en Configuration (bloc Indirect), comme dans le Delphi d'origine.
            double? fNominale = MesureConfig.ModeMesure == ModeMesure.Indirect
                ? MesureConfig.FNominale
                : (double?)null;

            if (nbRepetitions > 1)
                await LancerSerieStabAsync(MesureConfig, rubi, fNominale, nbRepetitions);
            else
                await LancerMesureAsync(MesureConfig, rubi, fNominale, preambule: "▶ Lancement");
        }

        /// <summary>
        /// Enchaîne <paramref name="nbRepetitions"/> stabs identiques, chacune après la fin RÉELLE de la
        /// précédente (l'estimation ne sert qu'à afficher l'heure de fin). Chaque stab crée son fichier
        /// (stab, stab1…, cf. ExcelService.InitialiserRapportAsync). Un arrêt, une erreur ou un hors-module interrompt la série.
        /// </summary>
        private async Task LancerSerieStabAsync(Mesure config, Rubidium rubi, double? fNominale, int nbRepetitions)
        {
            _serieStabEnCours = true;
            var swSerie = System.Diagnostics.Stopwatch.StartNew();
            var estimation = EstimationDureeStab.EstimerBalayage(config.GateIndices, config.NbMesures);
            int faites = 0;
            try
            {
                var totalEstime = TimeSpan.FromTicks(estimation.Ticks * nbRepetitions);
                Log("═══════════════════════════════════════════");
                Log($"🔁 Série de {nbRepetitions} stabs · ≈ {EstimationDureeStab.Formater(estimation)} par stab"
                  + $" · fin estimée vers {EstimationDureeStab.HeureFin(totalEstime)}");
                Journal.Info(CategorieLog.Mesure, "STAB_SERIE_DEBUT",
                    $"Série de {nbRepetitions} stabs, FI {config.NumFI}, gates {string.Join(",", config.GateIndices)}, "
                  + $"durée estimée {EstimationDureeStab.Formater(estimation)} par stab.");
                JournalFIService.NoterSerie(
                    $"{nbRepetitions} stabs · ≈ {EstimationDureeStab.Formater(estimation)} par stab");

                for (int r = 1; r <= nbRepetitions; r++)
                {
                    // Dès qu'une stab est finie, on remplace l'estimation par la durée réelle moyenne.
                    var parStab = faites > 0 ? TimeSpan.FromTicks(swSerie.Elapsed.Ticks / faites) : estimation;
                    var reste = TimeSpan.FromTicks(parStab.Ticks * (nbRepetitions - r + 1));
                    if (r > 1)
                        Log($"🔁 Stab {r}/{nbRepetitions} · fin de série estimée vers {EstimationDureeStab.HeureFin(reste)}");

                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    _repetitionStab = r;
                    _nbRepetitionsStab = nbRepetitions;
                    bool ok = await LancerMesureAsync(config, rubi, fNominale,
                        preambule: $"▶ Stab {r}/{nbRepetitions}");
                    sw.Stop();

                    if (!ok)
                    {
                        Log($"✖ Série interrompue à la stab {r}/{nbRepetitions} : les suivantes ne sont pas lancées.");
                        Journal.Warn(CategorieLog.Mesure, "STAB_SERIE_INTERROMPUE",
                            $"Série interrompue à la stab {r}/{nbRepetitions} ({faites} terminée(s)).");
                        JournalFIService.NoterSerie(
                            $"interrompue à la stab {r}/{nbRepetitions} ({faites} terminée(s))");
                        break;
                    }

                    faites++;
                    Journal.Info(CategorieLog.Mesure, "STAB_SERIE_REPETITION_FIN",
                        $"Stab {r}/{nbRepetitions} terminée en {EstimationDureeStab.Formater(sw.Elapsed)} "
                      + $"(estimée {EstimationDureeStab.Formater(estimation)}).",
                        new { r, nbRepetitions, reelleS = sw.Elapsed.TotalSeconds, estimeeS = estimation.TotalSeconds });
                }

                if (faites == nbRepetitions)
                {
                    Log("═══════════════════════════════════════════");
                    Log($"✅ Série terminée : {faites} stabs en {EstimationDureeStab.Formater(swSerie.Elapsed)}.");
                    JournalFIService.NoterSerie(
                        $"terminée · {faites} stabs en {EstimationDureeStab.Formater(swSerie.Elapsed)}");
                }
            }
            finally
            {
                _serieStabEnCours = false;
                _repetitionStab = 1;
                _nbRepetitionsStab = 1;
            }
        }

        [RelayCommand(CanExecute = nameof(PeutRelancer))]
        private async Task RelancerMesureAsync()
        {
            if (MesureEnCours || _serieStabEnCours || !DerniereMesureDisponible) return;
            var rubi = EtatApplication.RubidiumActif;
            if (rubi == null)
            {
                MessageBox.Show("Le rubidium actif n'est plus défini.",
                    "Impossible de relancer", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Stabilité : chaque stab écrit dans son propre fichier (stab, stab1…), rien à « écraser » ;
            // on rejoue la même stab (mêmes gates, mêmes répétitions) sans toucher aux rapports existants.
            if (MesureConfig.TypeMesure == TypeMesure.Stabilite)
            {
                await RelancerStabiliteAsync(rubi);
                return;
            }

            // Oui = écraser (repart d'un fichier neuf depuis le template) ; Non = conserver (défaut,
            // feuille ajoutée au fichier existant) ; Annuler = rien.
            var choix = MessageBox.Show(
                "Voulez-vous écraser les mesures précédentes ?\n\n"
              + "• Oui : repart d'un fichier vierge (les anciennes mesures sont perdues).\n"
              + "• Non : conserve l'historique et ajoute la nouvelle mesure à la suite.\n"
              + "• Annuler : ne lance rien.",
                "Relancer la mesure",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (choix == MessageBoxResult.Cancel) return;

            if (choix == MessageBoxResult.Yes)
            {
                // Classeur ouvert dans Excel COM : réinitialisé SUR PLACE (feuilles freq*/stab* retirées,
                // Recap nettoyé) sans fermer le Workbook, pour éviter la fenêtre grise SDI qu'entraînent
                // FermerClasseurActif + File.Delete. Sinon (Excel fermé à la main), fermer + supprimer.
                if (ExcelInteropHost.Instance.AClasseurActif)
                {
                    try
                    {
                        await ExcelInteropHost.Instance.ReinitialiserClasseurActifAsync();
                        Log("🗑 Classeur réinitialisé in-place (anciennes feuilles supprimées, fichier conservé).");
                        Journal.Info(CategorieLog.Mesure, "MESURE_RELANCE_ECRASE",
                            "Classeur réinitialisé via COM (pas de fermeture, pas de shell).");
                    }
                    catch (Exception ex)
                    {
                        Log($"⚠ Réinitialisation in-place échouée : {ex.Message} — fallback fermer + delete.");
                        try { await ExcelInteropHost.Instance.FermerClasseurActifAsync(); } catch { }
                        try { SupprimerRapportTousFormats(_excelService.CheminFichierGenere); } catch { }
                    }
                }
                else
                {
                    // Aucun classeur ouvert : suppression directe du fichier.
                    string cible = _excelService.CheminFichierGenere;
                    if (!string.IsNullOrEmpty(cible)
                        && (System.IO.File.Exists(cible) || System.IO.File.Exists(System.IO.Path.ChangeExtension(cible, ".xls"))))
                    {
                        try
                        {
                            SupprimerRapportTousFormats(cible);
                            Log($"🗑 Fichier précédent supprimé : {System.IO.Path.GetFileNameWithoutExtension(cible)}");
                            Journal.Info(CategorieLog.Mesure, "MESURE_RELANCE_ECRASE",
                                $"Fichier supprimé avant relance : {cible}");
                        }
                        catch (Exception ex)
                        {
                            Log($"⚠ Impossible de supprimer le fichier précédent : {ex.Message}");
                            MessageBox.Show(
                                $"Impossible de supprimer le fichier précédent :\n{ex.Message}\n\n"
                              + "Ferme-le dans Excel et relance.",
                                "Fichier verrouillé", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }
                    }
                }
            }

            await LancerMesureAsync(MesureConfig, rubi, _derniereFNominale,
                preambule: choix == MessageBoxResult.Yes
                    ? "🔁 Relance (fichier neuf)"
                    : "🔁 Relance (mêmes paramètres)");
        }

        private bool PeutRelancer() => DerniereMesureDisponible && !MesureEnCours;

        /// <summary>« Relancer » en stabilité : confirmation (gates, nombre de stabs, durée estimée)
        /// puis même lancement que la dernière fois, chaque stab dans un nouveau fichier.</summary>
        private async Task RelancerStabiliteAsync(Rubidium rubi)
        {
            int nb = Math.Max(1, _derniereNbRepetitionsStab);
            var duree = TimeSpan.FromTicks(
                EstimationDureeStab.EstimerBalayage(MesureConfig.GateIndices, MesureConfig.NbMesures).Ticks * nb);

            var choix = MessageBox.Show(
                "Relancer la même stabilité ?\n\n"
              + $"• {DescriptionGatesStab(MesureConfig)} · {MesureConfig.NbMesures} mesures par gate\n"
              + $"• {(nb > 1 ? $"{nb} stabs enchaînées" : "1 stab")}"
              + $" · durée estimée ≈ {EstimationDureeStab.Formater(duree)} (fin vers {EstimationDureeStab.HeureFin(duree)})\n\n"
              + "Chaque stab est enregistrée dans un nouveau fichier : les rapports précédents ne sont pas modifiés.",
                "Relancer la stabilité", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (choix != MessageBoxResult.Yes) return;

            if (nb > 1)
                await LancerSerieStabAsync(MesureConfig, rubi, _derniereFNominale, nb);
            else
                await LancerMesureAsync(MesureConfig, rubi, _derniereFNominale, preambule: "🔁 Relance");
        }

        /// <summary>« toutes les gates », « gate 1 s » ou « gates 1 s, 10 s, 100 s ».</summary>
        private static string DescriptionGatesStab(Mesure config)
        {
            var gates = config.GateIndices.Distinct().OrderBy(i => i).ToList();
            return gates.Count >= EnTetesMesureHelper.LibellesCanoniques.Count
                ? "toutes les gates"
                : gates.Count == 1
                    ? $"gate {EnTetesMesureHelper.LibelleGate(gates[0])}"
                    : $"gates {string.Join(", ", gates.Select(EnTetesMesureHelper.LibelleGate))}";
        }

        /// <summary>
        /// Écrase le rapport déjà présent (<paramref name="cheminAttendu"/>) avant une nouvelle
        /// mesure sur une FI au numéro repris : réinitialisation IN-PLACE si c'est le classeur ouvert
        /// dans Excel (comme « Relancer »), sinon suppression sur disque. Renvoie false si la
        /// suppression échoue (fichier verrouillé), message déjà affiché.
        /// </summary>
        /// <summary>Premier module de la config (principal, puis Fréquence auxiliaire en tachy) qui
        /// n'est pas marqué « dégradé », ou null si tous le sont (ou si aucun n'est choisi).</summary>
        private static string? ModuleNonDegrade(Mesure config)
        {
            static bool EstDegrade(TypeMesure type, string num) =>
                Metrologo.Services.Incertitude.ModulesIncertitudeService.Lister(type).Any(m =>
                    m.Degrade && string.Equals(m.NumModule, num, StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrWhiteSpace(config.NumModuleIncertitude)
                && !EstDegrade(config.TypeMesure, config.NumModuleIncertitude))
                return config.NumModuleIncertitude;
            if (EnTetesMesureHelper.EstTachymetre(config.TypeMesure)
                && !string.IsNullOrWhiteSpace(config.NumModuleIncertitudeFreq)
                && !EstDegrade(TypeMesure.Frequence, config.NumModuleIncertitudeFreq))
                return config.NumModuleIncertitudeFreq;
            return null;
        }

        /// <summary>Supprime le rapport sous ses deux formats possibles (.xlsx de travail, .xls de
        /// sortie). Lève une exception si un fichier existant ne peut pas être supprimé.</summary>
        private static void SupprimerRapportTousFormats(string cheminXlsx)
        {
            if (string.IsNullOrEmpty(cheminXlsx)) return;
            foreach (var chemin in new[] { cheminXlsx, System.IO.Path.ChangeExtension(cheminXlsx, ".xls") })
                if (System.IO.File.Exists(chemin)) System.IO.File.Delete(chemin);
        }

        private async Task<bool> EcraserRapportExistantAsync(string cheminAttendu)
        {
            bool estClasseurActif = ExcelInteropHost.Instance.AClasseurActif
                && (MemeChemin(ExcelInteropHost.Instance.CheminClasseurActif, cheminAttendu)
                    || MemeChemin(ExcelInteropHost.Instance.CheminClasseurActif, System.IO.Path.ChangeExtension(cheminAttendu, ".xls")));

            if (estClasseurActif)
            {
                try
                {
                    await ExcelInteropHost.Instance.ReinitialiserClasseurActifAsync();
                    Log("🗑 Classeur réinitialisé in-place (anciennes feuilles supprimées, fichier conservé).");
                    Journal.Info(CategorieLog.Mesure, "RAPPORT_ECRASE_INPLACE",
                        "Rapport FI réinitialisé via COM avant nouvelle mesure (pas de shell grise).");
                    return true;
                }
                catch (Exception ex)
                {
                    Log($"⚠ Réinitialisation in-place échouée : {ex.Message} — fallback fermer + supprimer.");
                    try { await ExcelInteropHost.Instance.FermerClasseurActifAsync(); } catch { }
                    // on enchaîne sur la suppression disque ci-dessous
                }
            }

            try
            {
                SupprimerRapportTousFormats(cheminAttendu);
                Log($"🗑 Ancien rapport supprimé : {System.IO.Path.GetFileNameWithoutExtension(cheminAttendu)}");
                Journal.Info(CategorieLog.Mesure, "RAPPORT_ECRASE",
                    $"Ancien rapport supprimé avant nouvelle mesure : {cheminAttendu}");
                return true;
            }
            catch (Exception ex)
            {
                Log($"⚠ Impossible de supprimer l'ancien rapport : {ex.Message}");
                MessageBox.Show(
                    $"Impossible de supprimer l'ancien rapport :\n{cheminAttendu}\n\n{ex.Message}\n\n"
                  + "Il est peut-être encore ouvert dans Excel. Ferme-le puis relance.",
                    "Fichier verrouillé", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
        }

        private static bool MemeChemin(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return false;
            try
            {
                return string.Equals(System.IO.Path.GetFullPath(a),
                    System.IO.Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase);
            }
            catch { return string.Equals(a, b, StringComparison.OrdinalIgnoreCase); }
        }

        private static string ResolvedNomAppareil(Mesure config)
        {
            if (string.IsNullOrEmpty(config.IdModeleCatalogue)) return "(aucun)";
            var modele = CatalogueAppareilsService.Instance.Modeles
                .FirstOrDefault(m => m.Id == config.IdModeleCatalogue);
            return modele?.Nom ?? $"({config.IdModeleCatalogue})";
        }

        /// <returns>true si la mesure est allée au bout ; false si arrêt, échec ou erreur.</returns>
        private async Task<bool> LancerMesureAsync(Mesure config, Rubidium rubi, double? fNominale, string preambule)
        {
            bool succes = false;

            // Paillasse : modules dégradés uniquement. La fenêtre de configuration filtre déjà ;
            // ce garde-fou couvre « Relancer », qui réutilise la config sans la rouvrir.
            if (!EstSurBaie && ModuleNonDegrade(config) is string moduleRefuse)
            {
                Log($"✖ Mesure non lancée : le module {moduleRefuse} n'est pas dégradé (mode Paillasse).");
                MessageBox.Show(
                    $"En mode Paillasse, seuls les modules d'incertitude dégradés sont autorisés.\n\n"
                  + $"Le module « {moduleRefuse} » ne l'est pas : choisis un module dégradé dans la configuration.",
                    "Module non autorisé en Paillasse", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            // FI passée en local (réseau absent ou lent au départ) : retour sur le réseau dès qu'il
            // répond (rapatriement du dossier local, puis suppression).
            await CheminsMetrologo.ReevaluerDossierFIAsync(config.NumFI ?? string.Empty);

            // Un Excel externe (hors notre instance COM cachée) peut verrouiller le .xlsx de la mesure et
            // faire échouer ClosedXML : on distingue les « fantômes » (reliquats COM sans fenêtre)
            // des vrais classeurs ouverts.
            var (_, excelsFantomes) = ExcelInteropHost.Instance.ListerExcelsExternesClasses();

            // 1. Fantômes (sans fenêtre, jamais un document utilisateur) : fermés discrètement pour libérer un verrou.
            if (excelsFantomes.Count > 0)
            {
                int nbFantomes = ExcelInteropHost.Instance.FermerExcels(excelsFantomes);
                Journal.Info(CategorieLog.Excel, "EXCELS_FANTOMES_NETTOYES",
                    $"{nbFantomes} processus EXCEL.EXE résiduel(s) (sans fenêtre) fermé(s) "
                  + "automatiquement avant la mesure.");
            }

            // 2. Classeurs VRAIMENT ouverts par l'utilisateur : on ne tue pas l'instance Excel (ses fichiers
            //    perso non sauvegardés seraient perdus) ; on ferme seulement les classeurs de mesure
            //    (nom freq/stab) via la Running Object Table, ce qui libère le verrou sur le rapport.
            int nbMesureFermes = ExcelInteropHost.Instance.FermerClasseursMesureOuvertsExternes();
            if (nbMesureFermes > 0)
            {
                Log($"🗑 {nbMesureFermes} classeur(s) de mesure (freq/stab) déjà ouvert(s) fermé(s) avant lancement.");
                Journal.Info(CategorieLog.Excel, "EXCELS_MESURE_FERMES_PARNOM",
                    $"{nbMesureFermes} classeur(s) de mesure freq/stab fermé(s) par leur nom avant la mesure "
                  + "(aucun process Excel tué, fichiers utilisateur préservés).");
            }

            // 3. Rapport sorti en .xls par la mesure précédente (réglage Admin) : on le repasse en
            //    .xlsx, le seul format sur lequel le moteur des rapports sait travailler. Il
            //    ressortira en .xls à la fin de cette mesure. Stabilité non concernée.
            if (config.TypeMesure != TypeMesure.Stabilite)
            {
                string cheminRapport = await Task.Run(() => _excelService.CalculerCheminFichierAttendu(config));
                if (!await ExcelInteropHost.Instance.ReconvertirRapportEnXlsxAsync(cheminRapport))
                {
                    string nomXls = System.IO.Path.ChangeExtension(System.IO.Path.GetFileName(cheminRapport), ".xls");
                    Log($"✖ Mesure non lancée : impossible de reprendre {nomXls} (fichier ouvert ailleurs ?).");
                    MessageBoxPremierPlan.Afficher(
                        $"Le rapport {nomXls} de la FI {config.NumFI} n'a pas pu être repris "
                      + "(il est probablement ouvert dans un autre logiciel).\n\n"
                      + "Ferme-le, puis relance la mesure.",
                        "Rapport indisponible", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
            }

            _cts = new CancellationTokenSource();
            _abandonCts = new CancellationTokenSource();
            MesureEnCours = true;

            // Heure de fin estimée dans le bandeau rouge : stabilité seulement.
            if (config.TypeMesure == TypeMesure.Stabilite)
                SuiviFinStab.Demarrer(config.GateIndices, config.NbMesures, _repetitionStab, _nbRepetitionsStab);
            else
                SuiviFinStab.Arreter();

            // Fenêtre flottante « ARRÊTER LA MESURE » sur un THREAD UI À PART (STA + Dispatcher.Run) : le clic
            // est pris en compte même quand le thread principal est saturé par les Interop Excel COM.
            StopMesureFloatingWindow? winStop = null;
            Thread? threadStop = null;
            var winReady = new System.Threading.ManualResetEventSlim(false);

            try
            {
                threadStop = new Thread(() =>
                {
                    try
                    {
                        winStop = new StopMesureFloatingWindow();
                        winStop.Configurer(() =>
                        {
                            // === ACTIONS VITALES EXÉCUTÉES DIRECTEMENT SUR CE THREAD ===
                            // CRUCIAL : ne PAS marshaler vers le main dispatcher, probablement bloqué par les
                            // Interop Excel COM en cours (le BeginInvoke ne s'exécuterait jamais).
                            //   • CancellationTokenSource.Cancel() : thread-safe.
                            //   • AborterMesureEnCours() : Device Clear GPIB via son propre driver IEEE.
                            //   • TuerProcessExcelAsync() : kill EXCEL.EXE par PID SANS prendre le lock _sync,
                            //     ce qui libère les COM calls bloqués (Excel fermé à la main pendant une
                            //     écriture : le RPC resterait bloqué ~60 s).
                            try { _cts?.Cancel(); } catch { /* swallow */ }
                            try { _orchestrator.AborterMesureEnCours(); }
                            catch (Exception ex)
                            {
                                Journal.Warn(CategorieLog.Mesure, "MESURE_STOP_SDC_ERR",
                                    $"Device Clear d'arrêt échoué : {ex.Message} — annulation continue malgré tout.");
                            }
                            // Pas de FermerClasseurActifAsync : il prendrait le lock potentiellement bloqué.
                            try { _ = ExcelInteropHost.Instance.TuerProcessExcelAsync(); }
                            catch { /* swallow */ }

                            // SIGNAL D'ABANDON : débloque le Task.WhenAny de LancerMesureAsync pour que l'UI
                            // reprenne la main sans attendre l'orchestration (bloquée ~60 s par un COM mort).
                            try { _abandonCts?.Cancel(); } catch { /* swallow */ }

                            Journal.Warn(CategorieLog.Mesure, "MESURE_STOP_FLOATING",
                                "Arrêt demandé via le bouton flottant (thread flottant indépendant).");

                            // Effets UI best-effort : BeginInvoke ne bloque pas, affichés si le main dispatcher répond.
                            try
                            {
                                Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    ArretEnCours = true;
                                    Log("⏹ Arrêt demandé via bouton flottant — annulation propagée.");
                                }), System.Windows.Threading.DispatcherPriority.Send);
                            }
                            catch { /* swallow */ }
                        });
                        winStop.Closed += (_, _) =>
                            System.Windows.Threading.Dispatcher.CurrentDispatcher
                                .BeginInvokeShutdown(System.Windows.Threading.DispatcherPriority.Background);
                        winStop.Show();
                        winReady.Set();
                        System.Windows.Threading.Dispatcher.Run();
                    }
                    catch (Exception ex)
                    {
                        Journal.Warn(CategorieLog.Systeme, "STOP_FLOATING_OPEN_ERR",
                            $"Affichage bouton STOP flottant échoué : {ex.Message}");
                        winReady.Set();
                    }
                });
                threadStop.SetApartmentState(System.Threading.ApartmentState.STA);
                threadStop.IsBackground = true;
                threadStop.Start();
                // Attend la fenêtre 1 s max : la mesure ne doit pas être pénalisée.
                winReady.Wait(TimeSpan.FromSeconds(1));
            }
            catch (Exception ex)
            {
                Journal.Warn(CategorieLog.Systeme, "STOP_FLOATING_OPEN_ERR",
                    $"Thread STOP flottant non démarré : {ex.Message}");
            }

            // Le classeur Excel n'est volontairement PAS fermé ici : le fermer rendrait AClasseurActif faux, la
            // voie COM (ajout d'une feuille sans fermer) ne serait jamais empruntée et le fallback ClosedXML
            // provoquerait la fenêtre grise SDI. MesureOrchestrator.ExecuterAsync le ferme lui-même en voie
            // ClosedXML (mesure 1, changement de FI) ; en voie COM (mesure 2+, même FI) il reste ouvert.

            string nomAppareilLog = ResolvedNomAppareil(config);
            Log("═══════════════════════════════════════════");
            Log($"{preambule} : {config.NbMesures} mesures sur {nomAppareilLog}");
            Log($"   FI {config.NumFI} · Rubidium : {rubi.Designation}");

            Journal.Info(CategorieLog.Mesure, "MESURE_DEBUT",
                $"{preambule} : {config.NbMesures} mesures sur {nomAppareilLog} pour FI {config.NumFI}",
                new
                {
                    numFI = config.NumFI,
                    type = config.TypeMesure.ToString(),
                    appareil = nomAppareilLog,
                    idModele = config.IdModeleCatalogue,
                    nbMesures = config.NbMesures,
                    mode = config.ModeMesure.ToString(),
                    source = config.SourceMesure.ToString(),
                    gateIndex = config.GateIndex,
                    fNominale,
                    rubidium = rubi.Designation,
                    gps = rubi.AvecGPS
                });

            // Journal FI ouvert et rattaché à l'OPÉRATEUR COURANT avant d'écrire (couvre « Relancer » sans
            // reconfiguration et un changement d'utilisateur en cours de session : DemarrerSession
            // clôt l'ancien bloc et en ouvre un nouveau si l'opérateur a changé, sinon no-op).
            string utilisateurCourant = EtatApplication.UtilisateurConnecte?.NomComplet
                ?? EtatApplication.UtilisateurConnecte?.Login
                ?? "(inconnu)";
            JournalFIService.DemarrerSession(config.NumFI ?? string.Empty,
                utilisateurCourant, EstSurBaie ? "Baie" : "Paillasse");

            // Journal FI : la config (écrite ici, une fois les gates d'une stab choisies, et
            // seulement si elle a changé) ; la mesure elle-même part sur la ligne du résultat.
            JournalFIService.NoterConfiguration(EnTetesMesureHelper.LibelleType(config.TypeMesure),
                ResumeConfigJournalFI(config, nomAppareilLog));
            JournalFIService.DebutMesure(preambule);

            var progress = new Progress<ProgressionMesure>(p =>
            {
                if (p.DerniereValeur.HasValue)
                    Log($"   {p.Message} : {p.DerniereValeur.Value:F6} Hz");
                else
                    Log($"… {p.Message}");
            });

            try
            {
                // La mesure est attendue via Task.WhenAny avec taskAbandon (bouton STOP flottant) et un
                // watchdog Excel : si Excel meurt ou si STOP est forcé, l'UI reprend la main immédiatement
                // sans attendre la fin des COM hangs (~60 s).
                var taskMesure = _orchestrator.ExecuterAsync(
                    config, rubi, fNominale, progress, _cts.Token);

                var taskAbandon = Task.Run(async () =>
                {
                    try { await Task.Delay(Timeout.Infinite, _abandonCts!.Token); }
                    catch (OperationCanceledException) { /* abandonné, c'est normal */ }
                });

                // Watchdog : poll Excel chaque seconde ; s'il meurt pendant la mesure (fenêtre fermée,
                // crash…), annule tout et signale l'erreur.
                bool excelMortDetecte = false;
                var taskWatchdogExcel = Task.Run(async () =>
                {
                    // Tokens copiés en local (struct) : restent valides si le finally met _abandonCts/_cts
                    // à null après Dispose (sinon NullReferenceException).
                    var abandonToken = _abandonCts!.Token;
                    var mesureToken = _cts!.Token;

                    // Délai initial : laisse Excel démarrer à la 1re écriture (_excelPid pas encore connu).
                    await Task.Delay(2000, abandonToken).ContinueWith(_ => { });

                    // Il faut 3 constats « mort » consécutifs (≈ 3 s) avant d'abandonner : un
                    // redémarrage d'Excel en cours de mesure (instance perdue entre deux mesures)
                    // ne doit pas faire abandonner la mesure sur un seul relevé.
                    int constatsMort = 0;
                    while (!abandonToken.IsCancellationRequested
                        && !mesureToken.IsCancellationRequested)
                    {
                        try { await Task.Delay(1000, abandonToken); }
                        catch (OperationCanceledException) { break; }

                        // Ne surveille que si Excel a été démarré (PID connu).
                        if (!ExcelInteropHost.Instance.EstDemarre) { constatsMort = 0; continue; }

                        if (ExcelInteropHost.Instance.EstProcessExcelEnVie()) { constatsMort = 0; continue; }
                        if (++constatsMort >= 3)
                        {
                            excelMortDetecte = true;
                            Journal.Erreur(CategorieLog.Excel, "EXCEL_FERME_BRUTAL",
                                "Process Excel détecté mort pendant la mesure (3 relevés consécutifs) — abandon auto déclenché.");

                            // Cascade d'annulation : token + GPIB + signal abandon.
                            try { _cts?.Cancel(); } catch { }
                            try { _orchestrator.AborterMesureEnCours(); } catch { }
                            try { _abandonCts?.Cancel(); } catch { }
                            break;
                        }
                    }
                });

                var premiere = await Task.WhenAny(taskMesure, taskAbandon);

                if (premiere == taskAbandon)
                {
                    if (excelMortDetecte)
                    {
                        // Excel fermé brutalement : log panneau + journal, sans MessageBox bloquante
                        // (l'UI revient simplement à l'état « prêt »).
                        Log("═══════════════════════════════════════════");
                        Log("⚠ Excel a été fermé pendant la mesure.");
                        Log("✖ Mesure interrompue — rapport non valide.");
                        Log("ℹ Tu peux relancer une nouvelle configuration.");
                        Journal.Warn(CategorieLog.Mesure, "MESURE_ABANDON_EXCEL_MORT",
                            "Mesure interrompue automatiquement (Excel fermé brutalement).");
                        JournalFIService.NoterResultat(false, "Excel fermé pendant la mesure");
                    }
                    else
                    {
                        Log("═══════════════════════════════════════════");
                        Log("⏹ Mesure arrêtée par l'utilisateur (bouton STOP).");
                        Log("ℹ Tu peux relancer une nouvelle configuration.");
                        Journal.Warn(CategorieLog.Mesure, "MESURE_ABANDON",
                            "Mesure abandonnée par l'utilisateur (taskMesure orpheline en arrière-plan).");
                        JournalFIService.NoterResultat(false, "arrêtée par l'opérateur");
                    }
                    return false;
                }

                var result = await taskMesure;

                if (result.Succes)
                {
                    succes = true;
                    Log("───────────────────────────────────────────");
                    Log($"✅ Moyenne : {result.Moyenne:F6} Hz");
                    Log($"✅ Écart-type : {result.EcartType:E3} Hz");

                    // Fichier principal verrouillé par un Excel ouvert : un fichier timestampé a été créé, on le signale.
                    if (_excelService.FallbackTimestampUtilise)
                    {
                        Log($"⚠ Le fichier principal était ouvert dans Excel — résultats sauvegardés dans :");
                        Log($"   {System.IO.Path.GetFileName(_excelService.CheminFichierGenere)}");
                    }

                    Log($"✅ Rapport Excel ouvert.");

                    _derniereFNominale = fNominale;
                    DerniereMesureDisponible = true;

                    Journal.Info(CategorieLog.Mesure, "MESURE_FIN",
                        $"Mesure terminée : moyenne {result.Moyenne:F6} Hz, σ {result.EcartType:E3} Hz",
                        new { result.Moyenne, result.EcartType, nbValeurs = result.Valeurs.Count });

                    // Journal FI : qui a fait quoi et quand ; les valeurs, elles, sont dans le rapport Excel.
                    JournalFIService.NoterResultat(true, $"{result.Valeurs.Count} mesures");

                    // Transfert du dossier FI vers le réseau échoué : le dossier reste en local et sera
                    // retransféré au prochain démarrage (cf. TransfertReseauService).
                    if (result.TransfertReseauOk == false)
                    {
                        Log("⚠ Réseau injoignable — la FI est enregistrée en local (Bureau\\Metrologo).");
                        // Une seule fois par FI et par session : pas de pop-up à chaque feuille de la série.
                        if (_fiAvertiesLocal.Add(config.NumFI ?? string.Empty))
                        {
                            // Premier plan forcé : Excel (rapport) est devant à ce moment-là.
                            MessageBoxPremierPlan.Afficher(
                                $"Le partage réseau ({CheminsMetrologo.MesuresLocal}) est injoignable : "
                              + $"la FI {config.NumFI} est enregistrée sur ton Bureau "
                              + $"({CheminsMetrologo.DossierFILocal(config.NumFI ?? string.Empty)}).\n\n"
                              + "Causes possibles :\n"
                              + "  • Le lecteur réseau est temporairement indisponible\n"
                              + "  • Latence ou perte de connexion\n"
                              + "  • Droits insuffisants sur le dossier cible\n\n"
                              + "Rien n'est perdu. Dès que le réseau revient (au prochain démarrage de "
                              + "Metrologo), le dossier y est rapatrié automatiquement puis supprimé du "
                              + "Bureau. D'ici là, le logiciel extérieur ne voit pas ces mesures.",
                                "Enregistrement en local",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }

                    // Sortie en .xls demandée (réglage Admin) mais ratée : le rapport est resté en
                    // .xlsx, que le logiciel extérieur ne lit pas encore.
                    if (result.ConversionXlsOk == false)
                    {
                        Log("⚠ Rapport resté en .xlsx : l'enregistrement en .xls a échoué (fichier .xls ouvert ailleurs ?).");
                        MessageBoxPremierPlan.Afficher(
                            $"La mesure est bien enregistrée, mais le rapport de la FI {config.NumFI} n'a pas pu "
                          + "être enregistré en .xls (le fichier .xls est probablement ouvert dans un autre logiciel).\n\n"
                          + "Il est resté en .xlsx : ferme le .xls, il sera reconverti à la prochaine mesure sur cette FI.",
                            "Rapport resté en .xlsx", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }

                    // Saisie post-mesure : fréquence lue + incertitudes — uniquement pour
                    // Fréquence + Source=Fréquencemètre (conforme Delphi : skip si Générateur).
                    await SaisiePostMesureAsync(config);
                }
                else
                {
                    Log($"✖ Échec : {result.Erreur}");
                    Journal.Erreur(CategorieLog.Mesure, "MESURE_ECHEC", result.Erreur ?? "Échec inconnu.");
                    JournalFIService.NoterResultat(false, result.Erreur ?? "Échec inconnu");
                    // Premier plan forcé : Excel est souvent devant, la pop-up doit passer PAR-DESSUS le rapport.
                    MessageBoxPremierPlan.Afficher(result.Erreur ?? "Erreur inconnue.",
                        "Mesure interrompue", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                Log($"✖ Erreur inattendue : {ex.Message}");
                Journal.Erreur(CategorieLog.Mesure, "MESURE_EXCEPTION", ex.Message, new { ex.StackTrace });
                JournalFIService.NoterResultat(false, $"erreur inattendue : {ex.Message}");
                MessageBoxPremierPlan.Afficher(ex.Message, "Erreur inattendue",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SuiviFinStab.Arreter();
                MesureEnCours = false;
                ArretEnCours = false;
                _cts?.Dispose();
                _cts = null;
                _abandonCts?.Dispose();
                _abandonCts = null;

                // Fermeture de la fenêtre flottante STOP via SON Dispatcher (thread séparé).
                try
                {
                    if (winStop != null)
                    {
                        winStop.Dispatcher.Invoke(() =>
                        {
                            try { winStop.Close(); } catch { /* swallow */ }
                        });
                    }
                }
                catch { /* swallow — pas critique si la fermeture rate */ }
            }

            return succes;
        }

        /// <summary>
        /// Flux post-mesure : dispatch selon le type de mesure :
        /// <list type="bullet">
        ///   <item>Fréquence + Fréquencemètre → page unique fréq. lue + résolution + incert. sup.</item>
        ///   <item>Tachy Contact/Optique et Stroboscope → page minimale résolution (tr/min) uniquement.</item>
        ///   <item>Autres types → pas de popup.</item>
        /// </list>
        /// Les valeurs saisies sont injectées dans les zones nommées du classeur via Interop.
        /// </summary>
        private async Task SaisiePostMesureAsync(Mesure config)
        {
            if (config.TypeMesure == TypeMesure.Frequence
                && config.SourceMesure == SourceMesure.Frequencemetre)
            {
                await SaisiePostMesureFrequenceAsync(config);
                return;
            }

            if (EnTetesMesureHelper.EstTachymetre(config.TypeMesure))
            {
                await SaisiePostMesureTachyAsync(config);
                return;
            }

            // Intervalle de temps : pas de saisie, mais le rapport est SAUVEGARDÉ comme en fréquence. Sinon le
            // classeur reste modifié (Saved=false) et sa fermeture à la mesure suivante (Workbook.Close
            // SaveChanges=true) doit le sauver : transition grise, risque de gel sur un dialogue Excel caché.
            if (config.TypeMesure == TypeMesure.Interval)
            {
                await ExcelInteropHost.Instance.SauvegarderAsync();
            }
        }

        private async Task SaisiePostMesureFrequenceAsync(Mesure config)
        {
            var vm = new SaisiePostMesureFreqViewModel(
                config.FNominale, config.Resolution, config.IncertSupp);

            var win = new SaisiePostMesureFreqWindow(vm) { Owner = Application.Current.MainWindow };
            if (win.ShowDialog() != true)
            {
                Log("ℹ Saisie post-mesure ignorée.");
                return;
            }

            config.Resolution = vm.Resolution;
            config.IncertSupp = vm.IncertSupp;

            Log($"📝 Post-mesure : FLue={vm.FrequenceLue:F6} Hz · Résolution={vm.Resolution:E3} · IncertSupRel={vm.IncertSupp:E3}");

            // La fréquence saisie est une valeur d'affichage : colonne 5 « fréquence indiquée » du récap
            // uniquement (comme « Géné. » en mode Générateur), pas dans la feuille de mesure. ZNFreqRef
            // reste à la valeur du rubidium et ZNFreqCorr n'en dépend pas.
            await ExcelInteropHost.Instance.EcrireFreqIndiqueeRecapAsync(vm.FrequenceLue);

            // Les incertitudes restent des zones nommées de la feuille (formules du récap).
            await ExcelInteropHost.Instance.EcrireZoneNommeeAsync("ZNIncertResol", vm.Resolution);
            await ExcelInteropHost.Instance.EcrireZoneNommeeAsync("ZNIncertSup", vm.IncertSupp);
            await ExcelInteropHost.Instance.SauvegarderAsync();
        }

        private async Task SaisiePostMesureTachyAsync(Mesure config)
        {
            // Résolution saisie en tr/min, écrite dans deux zones : ZNIncertResolRpm (I25, brute, bloc stats
            // RPM du template tachy) et ZNIncertResol (F25, convertie en Hz pour les formules historiques).
            double initialeRpm = config.Resolution * 60.0;
            var vm = new SaisiePostMesureTachyViewModel(initialeRpm);

            var win = new SaisiePostMesureTachyWindow(vm) { Owner = Application.Current.MainWindow };
            if (win.ShowDialog() != true)
            {
                Log("ℹ Saisie post-mesure tachy ignorée.");
                return;
            }

            config.Resolution = vm.ResolutionHz;
            // Pas d'IncertSupp en tachy : dégradation portée par les coeffs C/D du module RPM.
            config.IncertSupp = 0.0;

            Log($"📝 Post-mesure tachy : Résolution={vm.ResolutionRpm:F4} tr/min ({vm.ResolutionHz:E3} Hz)");

            await ExcelInteropHost.Instance.EcrireZoneNommeeAsync("ZNIncertResolRpm", vm.ResolutionRpm);
            await ExcelInteropHost.Instance.EcrireZoneNommeeAsync("ZNIncertResol", vm.ResolutionHz);
            await ExcelInteropHost.Instance.EcrireZoneNommeeAsync("ZNIncertSup", 0.0);
            await ExcelInteropHost.Instance.SauvegarderAsync();
        }

        [RelayCommand(CanExecute = nameof(PeutStopper))]
        private void StopperMesure()
        {
            if (_cts != null && !_cts.IsCancellationRequested)
            {
                // Retour visuel immédiat, sans attendre que le compteur libère son :FETCh?.
                ArretEnCours = true;

                _cts.Cancel();
                try { _orchestrator.AborterMesureEnCours(); }
                catch (Exception ex)
                {
                    Journal.Warn(CategorieLog.Mesure, "MESURE_STOP_SDC_ERR",
                        $"Device Clear d'arrêt échoué : {ex.Message} — annulation continue malgré tout.");
                }
                Log("⏹ Arrêt demandé — en cours…");
                Journal.Warn(CategorieLog.Mesure, "MESURE_STOP", "Arrêt demandé par l'utilisateur.");
                JournalFIService.SignalerArret();
            }
        }

        private bool PeutStopper() => MesureEnCours && !ArretEnCours;

        /// <summary>
        /// Config telle qu'écrite dans le journal FI (le type est dans le séparateur qui la précède) :
        /// nombre de mesures, appareil et gate(s) (en stabilité : gates balayées ; mode et source sans
        /// sens). Ailleurs, le mode n'apparaît que s'il est indirect.
        /// </summary>
        private static string ResumeConfigJournalFI(Mesure config, string nomAppareil)
        {
            bool stab = config.TypeMesure == TypeMesure.Stabilite;
            string resume = $"{config.NbMesures} mesures" + (stab ? " par gate" : string.Empty);
            if (!string.IsNullOrWhiteSpace(nomAppareil)) resume += $" · {nomAppareil}";

            if (stab) return resume + $" · {DescriptionGatesStab(config)}";

            resume += $" · gate {EnTetesMesureHelper.LibelleGate(config.GateIndex)}";
            if (config.ModeMesure == ModeMesure.Indirect) resume += " · indirect";
            return resume;
        }

        // -------- Utilitaires --------

        // Plafond de lignes de InformationsGenerales : une string trop grosse ralentit le re-render de la
        // TextBox liée (200-500 ms par update, jusqu'à bloquer le bouton Arrêter).
        private const int LOG_MAX_LIGNES = 200;

        private void Log(string message)
        {
            string nouvelle = $"\n[{DateTime.Now:HH:mm:ss}] {message}";
            string courant = InformationsGenerales + nouvelle;

            // Au-delà du plafond, on garde la queue (lignes les plus récentes).
            int nbLignes = courant.Count(c => c == '\n');
            if (nbLignes > LOG_MAX_LIGNES)
            {
                int aRetirer = nbLignes - LOG_MAX_LIGNES;
                int idx = 0;
                for (int i = 0; i < aRetirer; i++)
                {
                    idx = courant.IndexOf('\n', idx) + 1;
                }
                courant = courant.Substring(idx);
            }

            InformationsGenerales = courant;
        }
    }
}
