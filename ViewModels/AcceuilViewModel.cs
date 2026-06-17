using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Services.Besancon;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Ieee;
using Metrologo.Services.Journal;
using Metrologo.Views;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace Metrologo.ViewModels
{
    public partial class AccueilViewModel : ObservableObject
    {
        // Driver IEEE bas niveau : VISA reel via NI-VISA (donc materiel branche).
        // Pour tourner sans materiel, basculer sur new SimulationIeeeDriver().
        // On a essaye Ni488Driver(0) (P/Invoke ni4882.dll) : memes perfs sur le 53131A,
        // la latence vient de l'instrument lui-meme, pas de la stack.
        private readonly IIeeeDriver _ieeeDriver = new VisaIeeeDriver(gpibBoard: 0);
        // Type concret (et pas IExcelService) pour pouvoir lire FallbackTimestampUtilise.
        private readonly ExcelService _excelService = new ExcelService();
        private readonly MesureOrchestrator _orchestrator;

        private CancellationTokenSource? _cts;

        /// <summary>Token d'abandon force : si on clique STOP alors que la tache est coincee sur un
        /// COM Excel mort (RPC ~60 s), il rend la main a l'UI sans attendre la fin de l'orchestration.</summary>
        private CancellationTokenSource? _abandonCts;

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

            // On se tient a jour si l'administrateur change le rubidium actif.
            EtatApplication.RubidiumActifChange += (_, _) =>
            {
                RubidiumActifTexte = EtatApplication.RubidiumActifTexte;
                _ = RafraichirBesanconAsync();   // rubidium change => on reevalue le suivi
            };

            // Idem si un scan est relance depuis Diagnostic GPIB (ajout d'un appareil).
            EtatApplication.AppareilsDetectesChange += (_, _) => RafraichirResumeScan();

            // On rafraichit le voyant et le rapport quand la tache Besancon vient de tourner
            // (recuperation quotidienne ou rattrapage d'une moyenne manquante).
            BesanconScheduler.StatutChange += (_, _) => _ = RafraichirBesanconAsync();

            // Scan GPIB initial en tache de fond : il ne doit pas bloquer l'ouverture de la fenetre.
            _ = Task.Run(ScannerInitialAsync);

            // Etat Besancon de depart.
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

                // Timeout pousse a 3s : certains appareils (le 53131A surtout) sont longs a repondre
                // au *IDN? juste apres une mise sous tension, ou quand le bus a ete perturbe.
                var resultats = await ScannerGpib.ScannerAsync(gpibBoard: 0, timeoutMs: 3000);

                // On met a jour l'etat global des appareils detectes (lu ensuite par Configuration, etc.).
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

            // Avertissement : une reponse *IDN? incoherente trahit souvent un conflit d'adresse
            // (deux appareils regles sur la meme adresse GPIB).
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
            // On privilegie le chemin reseau configure (Admin > Chemins d'acces > Mesures), et on
            // retombe sur Bureau\Metrologo si le reseau est vide ou injoignable
            // (connexion coupee, partage indisponible).
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
                // L'affectation ci-dessus reaffecte souvent la meme reference (le vm avait recu
                // notre instance) : du coup CommunityToolkit ne notifie pas et le bandeau
                // « Fiche en cours » ne s'affiche jamais. On force donc la notification.
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

                // Journal utilisateur FI : une entree detaillee avec les parametres cles.
                string gateLib = Metrologo.Services.EnTetesMesureHelper.LibelleGate(MesureConfig.GateIndex);
                JournalFIService.Ecrire("CONFIG_VALIDEE",
                    $"{MesureConfig.TypeMesure} · {MesureConfig.NbMesures} mesures · "
                    + $"{nomAppareil} · gate {gateLib} · {MesureConfig.ModeMesure} · {MesureConfig.SourceMesure}");

                // Enchainement automatique : une fois la config validee, on lance directement
                // la mesure (ouvrirConfigAvant=false pour ne pas rouvrir la fenetre de config en boucle).
                await ExecuterMesureInterneAsync(ouvrirConfigAvant: false);
            }
        }

        /// <summary>Bouton « Lancer la mesure » : on repasse toujours par la fenetre de config (l'utilisateur
        /// reconfigure a chaque mesure), qui rappelle ExecuterMesureInterneAsync(false) une fois validee.
        /// Pour relancer sans reconfigurer, c'est le bouton « Relancer ».</summary>
        [RelayCommand]
        private Task ExecuterMesureAsync() => ExecuterMesureInterneAsync(ouvrirConfigAvant: true);

        private async Task ExecuterMesureInterneAsync(bool ouvrirConfigAvant)
        {
            if (MesureEnCours) return;

            // 1) Le rubidium est obligatoire (et ne se definit que dans le menu Administration).
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

            // 2) Clic sur « Lancer la mesure » : on ouvre la config meme si une config valide
            //    etait deja en place (N° FI different, parametres modifies...). A la validation,
            //    OuvrirConfigurationAsync rappellera ExecuterMesureInterneAsync(false) pour
            //    enchainer sur le lancement reel. On return ici pour ne pas appeler deux fois.
            if (ouvrirConfigAvant || string.IsNullOrWhiteSpace(MesureConfig?.NumFI))
            {
                Log("ℹ Ouverture de la configuration avant lancement.");
                await OuvrirConfigurationAsync();
                return;
            }

            // 2bis) Le module d'incertitude est obligatoire : sans module choisi, les
            // coefficients A/B (et C/D en tachy) restent a des valeurs par defaut codees en dur,
            // ce qui fausse tout le calcul d'incertitude du rapport. On refuse donc le lancement
            // et on rouvre la fenetre Configuration pour que l'utilisateur fasse son choix.
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

            // 2ter) En tachymetrie, le module Frequence auxiliaire (pour ZNCoeffA/B en Hz) est
            // lui aussi obligatoire : il caracterise le frequencemetre support, donc sans lui le
            // rapport est incomplet cote Hz.
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

            // 3) Gate : deja choisie dans la fenetre Configuration (MesureConfig.GateIndex).
            //    Pour la Stabilite, on ouvre la fenetre dediee ou l'utilisateur selectionne les
            //    gates a balayer (une ou plusieurs, via cases a cocher + presets).
            if (MesureConfig.TypeMesure == TypeMesure.Stabilite)
            {
                var gateWin = new SelectionGateWindow(MesureConfig) { Owner = Application.Current.MainWindow };
                if (gateWin.ShowDialog() != true) { Log("✖ Mesure annulée (sélection gates)."); return; }
                MesureConfig.GateIndices = gateWin.ViewModel.IndicesGatesResultats;
            }
            else if (MesureConfig.GateIndices.Count > 1)
            {
                // Filet de securite : si on lance une mesure non-Stab juste apres une Stab, les
                // gates multiples choisies par la Stab trainent encore en memoire. On reduit la
                // liste a un seul element (la 1ere gate), sinon la mesure Frequence boucle sur
                // toutes les gates de la Stab precedente.
                MesureConfig.GateIndex = MesureConfig.GateIndices[0]; // le setter ramene la liste a 1 element
            }
            Log($"⏱ Gates à balayer : {string.Join(", ", MesureConfig.GateIndices)}");

            // 4) La frequence nominale est deja saisie dans ConfigurationWindow (bloc Indirect),
            //    comme dans le Delphi d'origine (pas de dialog pre-mesure pour FNominale).
            double? fNominale = MesureConfig.ModeMesure == ModeMesure.Indirect
                ? MesureConfig.FNominale
                : (double?)null;

            await LancerMesureAsync(MesureConfig, rubi, fNominale, preambule: "▶ Lancement");
        }

        [RelayCommand(CanExecute = nameof(PeutRelancer))]
        private async Task RelancerMesureAsync()
        {
            if (MesureEnCours || !DerniereMesureDisponible) return;
            var rubi = EtatApplication.RubidiumActif;
            if (rubi == null)
            {
                MessageBox.Show("Le rubidium actif n'est plus défini.",
                    "Impossible de relancer", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Deux possibilites : conserver les anciennes mesures (= ajouter une feuille au
            // fichier existant) ou ecraser (= repartir d'un fichier neuf depuis le template).
            //   Oui = ecraser (fichier supprime)
            //   Non = conserver (comportement par defaut, mesures ajoutees a la suite)
            //   Annuler = on ne lance rien.
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
                // OPTION A v2 : si un classeur est ouvert dans Excel COM, on le reinitialise SUR
                // PLACE (on enleve les feuilles freq*/stab*/etc. et on nettoie le Recap) SANS jamais
                // fermer le Workbook. Ca evite la fenetre grise SDI qui surgissait avec l'ancien flux
                // FermerClasseurActif + File.Delete (~1-3s a 0 Workbook = shell garanti).
                //
                // Si aucun classeur n'est ouvert (rare : l'utilisateur a ferme Excel a la main apres
                // la mesure precedente), on retombe sur l'ancien flux fermer + delete.
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
                        string cibleF = _excelService.CheminFichierGenere;
                        if (!string.IsNullOrEmpty(cibleF) && System.IO.File.Exists(cibleF))
                        {
                            try { System.IO.File.Delete(cibleF); } catch { }
                        }
                    }
                }
                else
                {
                    // Aucun classeur ouvert : ancien flux (rien a contourner, pas de shell ici).
                    string cible = _excelService.CheminFichierGenere;
                    if (!string.IsNullOrEmpty(cible) && System.IO.File.Exists(cible))
                    {
                        try
                        {
                            System.IO.File.Delete(cible);
                            Log($"🗑 Fichier précédent supprimé : {System.IO.Path.GetFileName(cible)}");
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

            JournalFIService.Ecrire("RELANCE",
                choix == MessageBoxResult.Yes
                    ? "fichier neuf (ancien supprimé)"
                    : "conserve historique");

            await LancerMesureAsync(MesureConfig, rubi, _derniereFNominale,
                preambule: choix == MessageBoxResult.Yes
                    ? "🔁 Relance (fichier neuf)"
                    : "🔁 Relance (mêmes paramètres)");
        }

        private bool PeutRelancer() => DerniereMesureDisponible && !MesureEnCours;

        private static string ResolvedNomAppareil(Mesure config)
        {
            if (string.IsNullOrEmpty(config.IdModeleCatalogue)) return "(aucun)";
            var modele = CatalogueAppareilsService.Instance.Modeles
                .FirstOrDefault(m => m.Id == config.IdModeleCatalogue);
            return modele?.Nom ?? $"({config.IdModeleCatalogue})";
        }

        private async Task LancerMesureAsync(Mesure config, Rubidium rubi, double? fNominale, string preambule)
        {
            // On regarde si un Excel externe (autre que notre instance COM cachee) est ouvert : il
            // peut garder verrouille le .xlsx de la mesure et faire echouer ClosedXML. On distingue
            // les reliquats COM sans fenetre (les « fantomes ») des vrais classeurs ouverts.
            var (_, excelsFantomes) = ExcelInteropHost.Instance.ListerExcelsExternesClasses();

            // 1. Fantomes (pas de fenetre = reliquat de pilotage COM, jamais un document de
            //    l'utilisateur) : on les ferme discretement pour liberer un eventuel verrou.
            if (excelsFantomes.Count > 0)
            {
                int nbFantomes = ExcelInteropHost.Instance.FermerExcels(excelsFantomes);
                Journal.Info(CategorieLog.Excel, "EXCELS_FANTOMES_NETTOYES",
                    $"{nbFantomes} processus EXCEL.EXE résiduel(s) (sans fenêtre) fermé(s) "
                  + "automatiquement avant la mesure.");
            }

            // 2. Classeurs VRAIMENT ouverts par l'utilisateur : on ne tue PLUS l'instance Excel
            //    (ca fermait aussi ses fichiers persos non sauvegardes). On ferme SEULEMENT les
            //    classeurs de mesure — ils portent toujours le nom freq/stab — par leur nom, via
            //    la Running Object Table. Tout autre fichier ouvert par l'utilisateur reste intact.
            //    Ca suffit a liberer le verrou sur le rapport de la FI.
            int nbMesureFermes = ExcelInteropHost.Instance.FermerClasseursMesureOuvertsExternes();
            if (nbMesureFermes > 0)
            {
                Log($"🗑 {nbMesureFermes} classeur(s) de mesure (freq/stab) déjà ouvert(s) fermé(s) avant lancement.");
                Journal.Info(CategorieLog.Excel, "EXCELS_MESURE_FERMES_PARNOM",
                    $"{nbMesureFermes} classeur(s) de mesure freq/stab fermé(s) par leur nom avant la mesure "
                  + "(aucun process Excel tué, fichiers utilisateur préservés).");
            }

            _cts = new CancellationTokenSource();
            _abandonCts = new CancellationTokenSource();
            MesureEnCours = true;

            // On affiche la fenetre flottante « ARRETER LA MESURE » sur un THREAD UI A PART
            // (STA + Dispatcher.Run). Comme ca le clic/raccourci est pris en compte tout de
            // suite, meme quand le thread principal est sature par les Interop Excel COM. Le
            // callback onStop repasse sur le Dispatcher principal pour derouler la sequence
            // d'arret (annulation CTS, abort GPIB, fermeture Excel).
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
                            // CRUCIAL : ne PAS marshaler ces appels vers le main dispatcher
                            // car celui-ci est probablement bloqué par les COM Interop Excel
                            // en cours (Workbooks.Open, Cell.Value2 = ..., etc.) — le
                            // BeginInvoke ne s'exécuterait jamais.
                            //   • CancellationTokenSource.Cancel() = thread-safe par design.
                            //   • Orchestrator.AborterMesureEnCours() envoie un Device Clear
                            //     GPIB qui utilise son propre driver IEEE indépendant.
                            //   • TuerProcessExcelAsync() = nucléaire : kill EXCEL.EXE par PID
                            //     SANS prendre le lock _sync → libère instantanément tous les
                            //     COM calls bloqués (cas typique : utilisateur a fermé Excel
                            //     manuellement pendant l'écriture d'une cellule, le RPC reste
                            //     bloqué ~60 s sinon).
                            try { _cts?.Cancel(); } catch { /* swallow */ }
                            try { _orchestrator.AborterMesureEnCours(); }
                            catch (Exception ex)
                            {
                                Journal.Warn(CategorieLog.Mesure, "MESURE_STOP_SDC_ERR",
                                    $"Device Clear d'arrêt échoué : {ex.Message} — annulation continue malgré tout.");
                            }
                            // Tue Excel directement — pas de FermerClasseurActifAsync qui prendrait
                            // le lock potentiellement bloqué. Le kill libère les COM hangs.
                            try { _ = ExcelInteropHost.Instance.TuerProcessExcelAsync(); }
                            catch { /* swallow */ }

                            // SIGNAL D'ABANDON : déclenche Task.WhenAny dans LancerMesureAsync
                            // pour que l'UI reprenne la main IMMÉDIATEMENT, sans attendre que
                            // la tâche orchestration finisse (qui peut être bloquée 60 s par
                            // un COM Excel mort). Critique pour ne pas avoir à relancer l'app.
                            try { _abandonCts?.Cancel(); } catch { /* swallow */ }

                            Journal.Warn(CategorieLog.Mesure, "MESURE_STOP_FLOATING",
                                "Arrêt demandé via le bouton flottant (thread flottant indépendant).");

                            // === Side-effects UI best-effort (peuvent rater si main thread bloqué) ===
                            // BeginInvoke ne bloque pas — si le main dispatcher répond plus tard,
                            // ces mises à jour s'afficheront ; sinon tant pis, l'essentiel est fait.
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
                // Bloque max 1 s pour la création de la fenêtre — si plus, on continue,
                // la mesure ne doit pas être pénalisée.
                winReady.Wait(TimeSpan.FromSeconds(1));
            }
            catch (Exception ex)
            {
                Journal.Warn(CategorieLog.Systeme, "STOP_FLOATING_OPEN_ERR",
                    $"Thread STOP flottant non démarré : {ex.Message}");
            }

            // OPTION A v2 : on NE FERME PLUS systématiquement le classeur Excel ici.
            //
            // C'était la cause profonde du bug "fenêtre grise" : à chaque clic sur Relancer
            // (ou Lancer la mesure), on faisait FermerClasseurActifAsync, ce qui mettait
            // _classeurActif à null → MesureOrchestrator évaluait AClasseurActif=False
            // → la voie COM v2 (qui ajoute une feuille sans fermer le classeur) n'était
            // JAMAIS empruntée → fallback ClosedXML → moment 0-Workbook = shell SDI grise.
            //
            // MesureOrchestrator.ExecuterAsync ferme déjà le classeur lui-même si la voie
            // ClosedXML est sélectionnée (cas mesure 1, ou changement de FI). En voie COM
            // (mesure 2+ même FI), on garde le classeur ouvert.

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

            // S'assure que le journal FI est ouvert ET rattaché à l'OPÉRATEUR COURANT avant
            // d'écrire. Couvre la reprise sans reconfiguration (bouton « Relancer », qui ne
            // repasse pas par la fenêtre de config) et le changement d'utilisateur en cours de
            // session : si l'opérateur a changé depuis le dernier bloc, DemarrerSession clôt
            // l'ancien bloc et en ouvre un nouveau à son nom ; sinon c'est un no-op.
            string utilisateurCourant = EtatApplication.UtilisateurConnecte?.NomComplet
                ?? EtatApplication.UtilisateurConnecte?.Login
                ?? "(inconnu)";
            JournalFIService.DemarrerSession(config.NumFI ?? string.Empty,
                utilisateurCourant, EstSurBaie ? "Baie" : "Paillasse");

            // Journal utilisateur FI : ligne MESURE_DEBUT avec préambule (Lancement / Relance).
            JournalFIService.Ecrire("MESURE_DEBUT",
                $"{preambule} · {config.NbMesures} mesures · {nomAppareilLog}");

            var progress = new Progress<ProgressionMesure>(p =>
            {
                if (p.DerniereValeur.HasValue)
                    Log($"   {p.Message} : {p.DerniereValeur.Value:F6} Hz");
                else
                    Log($"… {p.Message}");
            });

            try
            {
                // On lance la mesure mais on ne l'attend pas directement — on l'attend
                // dans un Task.WhenAny avec :
                //   - taskAbandon : signal manuel via le bouton STOP flottant
                //   - taskWatchdog : surveillance auto qui détecte si Excel meurt
                // Ainsi, si Excel est fermé brutalement (ou si l'utilisateur force STOP),
                // l'UI reprend la main IMMÉDIATEMENT sans attendre que les COM hangs
                // se résolvent (~60 s sinon).
                var taskMesure = _orchestrator.ExecuterAsync(
                    config, rubi, fNominale, progress, _cts.Token);

                var taskAbandon = Task.Run(async () =>
                {
                    try { await Task.Delay(Timeout.Infinite, _abandonCts!.Token); }
                    catch (OperationCanceledException) { /* abandonné, c'est normal */ }
                });

                // Watchdog : poll Excel toutes les secondes. Si Excel meurt pendant la
                // mesure (utilisateur a fermé la fenêtre Excel, crash, etc.), on annule
                // tout automatiquement + on affiche un message d'erreur.
                bool excelMortDetecte = false;
                var taskWatchdogExcel = Task.Run(async () =>
                {
                    // Capture des Token (struct value-type) en local : ils restent valides
                    // même si le finally a mis _abandonCts/_cts à null après Dispose pendant
                    // qu'on était dans l'await Task.Delay ci-dessous. Sans cette copie, on
                    // crashait en NullReferenceException sur _abandonCts.IsCancellationRequested.
                    var abandonToken = _abandonCts!.Token;
                    var mesureToken = _cts!.Token;

                    // Petit délai initial pour laisser le temps à Excel d'être démarré
                    // par la première écriture (cas où _excelPid n'est pas encore connu).
                    await Task.Delay(2000, abandonToken).ContinueWith(_ => { });

                    while (!abandonToken.IsCancellationRequested
                        && !mesureToken.IsCancellationRequested)
                    {
                        try { await Task.Delay(1000, abandonToken); }
                        catch (OperationCanceledException) { break; }

                        // Ne surveille que si Excel a été démarré (PID connu).
                        if (!ExcelInteropHost.Instance.EstDemarre) continue;

                        if (!ExcelInteropHost.Instance.EstProcessExcelEnVie())
                        {
                            excelMortDetecte = true;
                            Journal.Erreur(CategorieLog.Excel, "EXCEL_FERME_BRUTAL",
                                "Process Excel détecté mort pendant la mesure — abandon auto déclenché.");

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
                        // Excel fermé brutalement : on log dans le panneau + journal.
                        // Pas de MessageBox bloquante — l'utilisateur voit juste l'UI
                        // revenir en état "prêt" (boutons réactivés) avec un message
                        // d'info dans le panneau, exactement comme une fin de mesure
                        // normale mais avec un texte différent.
                        Log("═══════════════════════════════════════════");
                        Log("⚠ Excel a été fermé pendant la mesure.");
                        Log("✖ Mesure interrompue — rapport non valide.");
                        Log("ℹ Tu peux relancer une nouvelle configuration.");
                        Journal.Warn(CategorieLog.Mesure, "MESURE_ABANDON_EXCEL_MORT",
                            "Mesure interrompue automatiquement (Excel fermé brutalement).");
                    }
                    else
                    {
                        Log("═══════════════════════════════════════════");
                        Log("⏹ Mesure arrêtée par l'utilisateur (bouton STOP).");
                        Log("ℹ Tu peux relancer une nouvelle configuration.");
                        Journal.Warn(CategorieLog.Mesure, "MESURE_ABANDON",
                            "Mesure abandonnée par l'utilisateur (taskMesure orpheline en arrière-plan).");
                    }
                    return;
                }

                var result = await taskMesure;

                if (result.Succes)
                {
                    Log("───────────────────────────────────────────");
                    Log($"✅ Moyenne : {result.Moyenne:F6} Hz");
                    Log($"✅ Écart-type : {result.EcartType:E3} Hz");

                    // Si le fichier principal était verrouillé par un Excel ouvert, on a créé
                    // un fichier timestampé en parallèle — on le signale à l'utilisateur.
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

                    // Journal utilisateur FI : ligne MESURE_FIN avec stats clés.
                    JournalFIService.Ecrire("MESURE_FIN",
                        $"{result.Valeurs.Count} valeurs · moy {result.Moyenne:F6} Hz · σ {result.EcartType:E3} Hz");

                    // Si le transfert du dossier FI vers le réseau a échoué : prévenir l'utilisateur.
                    // Le dossier reste en local et sera retransféré automatiquement au prochain
                    // démarrage de Metrologo (cf. TransfertReseauService).
                    if (result.TransfertReseauOk == false)
                    {
                        Log("⚠ Transfert réseau échoué — le dossier FI reste en local.");
                        // Premier plan forcé : Excel (rapport) est devant à ce moment-là.
                        MessageBoxPremierPlan.Afficher(
                            $"Le dossier de la FI {config.NumFI} n'a pas pu être copié sur "
                          + $"le partage réseau ({CheminsMetrologo.MesuresLocal}).\n\n"
                          + "Causes possibles :\n"
                          + "  • Le lecteur réseau est temporairement indisponible\n"
                          + "  • Latence ou perte de connexion\n"
                          + "  • Droits insuffisants sur le dossier cible\n\n"
                          + "Toutes les données restent en local sur ton Bureau. Au prochain "
                          + "démarrage de Metrologo, le transfert sera retenté automatiquement.",
                            "Transfert réseau différé",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                    }

                    // Saisie post-mesure : fréquence lue + incertitudes — uniquement pour
                    // Fréquence + Source=Fréquencemètre (conforme Delphi : skip si Générateur).
                    await SaisiePostMesureAsync(config);
                }
                else
                {
                    Log($"✖ Échec : {result.Erreur}");
                    Journal.Erreur(CategorieLog.Mesure, "MESURE_ECHEC", result.Erreur ?? "Échec inconnu.");
                    JournalFIService.Ecrire("MESURE_ECHEC", result.Erreur ?? "Échec inconnu");
                    // Premier plan forcé : Excel est souvent devant (ex. valeur hors du
                    // module d'incertitude) — la pop-up doit passer PAR-DESSUS le rapport
                    // pour que l'utilisateur comprenne que la mesure n'est pas passée.
                    MessageBoxPremierPlan.Afficher(result.Erreur ?? "Erreur inconnue.",
                        "Mesure interrompue", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                Log($"✖ Erreur inattendue : {ex.Message}");
                Journal.Erreur(CategorieLog.Mesure, "MESURE_EXCEPTION", ex.Message, new { ex.StackTrace });
                JournalFIService.Ecrire("MESURE_ECHEC", $"Exception : {ex.Message}");
                MessageBoxPremierPlan.Afficher(ex.Message, "Erreur inattendue",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                MesureEnCours = false;
                ArretEnCours = false;
                _cts?.Dispose();
                _cts = null;
                _abandonCts?.Dispose();
                _abandonCts = null;

                // Ferme la fenêtre flottante STOP via SON Dispatcher (elle vit sur un thread
                // séparé, donc on doit lui demander gentiment de se fermer depuis là-bas).
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

            // Intervalle de temps : pas de saisie post-mesure, mais on SAUVEGARDE le rapport
            // comme le fait la fréquence. Sans ça le classeur reste "modifié" (Saved=false) :
            // au lancement de la mesure suivante, sa fermeture (FermerClasseurActifInterne →
            // Workbook.Close(SaveChanges=true), ExcelInteropHost.cs:3458) doit le sauver à ce
            // moment-là — ce qui provoque la transition grise entre deux mesures et peut figer
            // sur un dialogue Excel caché. En sauvant ici, la fermeture suivante est immédiate
            // et propre (mêmes procédures qu'en fréquence).
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

            // La fréquence saisie est une valeur d'affichage : elle va UNIQUEMENT dans la
            // colonne 5 « fréquence indiquée » du récap (même emplacement que « Géné. » en
            // mode Générateur). On ne l'écrit PLUS dans la feuille de mesure : la cellule
            // « Valeur de réf. (Hz) = » (ZNFreqRef) reste à la valeur du rubidium et n'est
            // donc plus polluée, et la fréquence corrigée (ZNFreqCorr) n'en dépend plus.
            await ExcelInteropHost.Instance.EcrireFreqIndiqueeRecapAsync(vm.FrequenceLue);

            // Les incertitudes restent des zones nommées de la feuille (formules du récap).
            await ExcelInteropHost.Instance.EcrireZoneNommeeAsync("ZNIncertResol", vm.Resolution);
            await ExcelInteropHost.Instance.EcrireZoneNommeeAsync("ZNIncertSup", vm.IncertSupp);
            await ExcelInteropHost.Instance.SauvegarderAsync();
        }

        private async Task SaisiePostMesureTachyAsync(Mesure config)
        {
            // Pour le tachy, l'utilisateur saisit la résolution en tr/min (unité naturelle
            // de l'appareil). On écrit deux zones :
            //   - ZNIncertResolRpm (cellule I25) : valeur brute saisie en tr/min, visible
            //     dans le bloc stats RPM du template tachy.
            //   - ZNIncertResol (cellule F25, Hz) : valeur convertie pour rester cohérente
            //     avec les formules historiques en Hz (incert. globale, etc.).
            double initialeRpm = config.Resolution * 60.0;
            var vm = new SaisiePostMesureTachyViewModel(initialeRpm);

            var win = new SaisiePostMesureTachyWindow(vm) { Owner = Application.Current.MainWindow };
            if (win.ShowDialog() != true)
            {
                Log("ℹ Saisie post-mesure tachy ignorée.");
                return;
            }

            config.Resolution = vm.ResolutionHz;
            // IncertSupp non saisie pour tachy : la dégradation est portée par les coeffs
            // C/D du module RPM, pas par une saisie manuelle.
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
                // Feedback visuel instantané : le bouton bascule en "Arrêt en cours…"
                // dès le clic, sans attendre que le compteur libère son :FETCh?.
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
                JournalFIService.Ecrire("STOP_UTILISATEUR", "Arrêt manuel via bouton flottant");
            }
        }

        private bool PeutStopper() => MesureEnCours && !ArretEnCours;

        // -------- Utilitaires --------

        // Plafond du nombre de lignes gardées dans InformationsGenerales pour éviter
        // que la string ne grossisse à plusieurs Ko, ce qui ralentit le re-render WPF
        // de la TextBox liée (chaque update prend alors 200-500 ms et bloque visuellement
        // l'UI — y compris le bouton Arrêter pendant une mesure longue).
        private const int LOG_MAX_LIGNES = 200;

        private void Log(string message)
        {
            string nouvelle = $"\n[{DateTime.Now:HH:mm:ss}] {message}";
            string courant = InformationsGenerales + nouvelle;

            // Trim au-delà du plafond : on garde la queue (les lignes les plus récentes).
            // Coût : O(n) une fois par overflow, négligeable face au gain UI.
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
