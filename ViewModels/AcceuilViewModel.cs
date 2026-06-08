using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Ieee;
using Metrologo.Services.Journal;
using Metrologo.Views;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace Metrologo.ViewModels
{
    public partial class AccueilViewModel : ObservableObject
    {
        // Driver IEEE bas niveau : VISA réel via NI-VISA (matériel branché).
        // Pour la simulation sans matériel, remplacer par `new SimulationIeeeDriver()`.
        // Pour P/Invoke direct sur ni4882.dll : `new Ni488Driver(0)` — testé, perf identique
        // sur 53131A car la latence est dans l'instrument lui-même, pas dans la stack.
        private readonly IIeeeDriver _ieeeDriver = new VisaIeeeDriver(gpibBoard: 0);
        // Type concret conservé (au lieu de IExcelService) pour accéder à FallbackTimestampUtilise.
        private readonly ExcelService _excelService = new ExcelService();
        private readonly MesureOrchestrator _orchestrator;

        private CancellationTokenSource? _cts;

        /// <summary>
        /// Token séparé pour signaler un ABANDON forcé de la tâche de mesure : utilisé
        /// quand l'utilisateur appuie sur STOP mais que la tâche est bloquée par un
        /// COM Excel mort (RPC qui prend ~60 s à se réveiller). Permet de rendre la main
        /// à l'UI immédiatement même si la tâche orchestration ne réagit pas.
        /// </summary>
        private CancellationTokenSource? _abandonCts;

        [ObservableProperty] private bool _estSurBaie = true;
        [ObservableProperty] private string _informationsGenerales = "Prêt. En attente d'exécution...";
        [ObservableProperty] private string _rubidiumActifTexte = EtatApplication.RubidiumActifTexte;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(EstConfigure))]
        [NotifyPropertyChangedFor(nameof(LibelleAppareilConfigure))]
        private Mesure _mesureConfig = new Mesure();

        /// <summary>
        /// Vrai dès qu'une configuration a été validée (N° FI renseigné). Pilote l'affichage
        /// du bandeau « Fiche en cours » sur l'écran d'accueil.
        /// </summary>
        public bool EstConfigure => !string.IsNullOrWhiteSpace(MesureConfig?.NumFI);

        /// <summary>
        /// Libellé court de l'appareil sélectionné dans la config (vide si aucun) — affiché
        /// dans le bandeau « Fiche en cours ».
        /// </summary>
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

        /// <summary>
        /// Vrai entre le clic sur « Arrêter » et la fin réelle de la mesure (token annulé +
        /// Device Clear envoyé). Donne un feedback visuel instantané au user (le bouton
        /// passe à « Arrêt en cours… ») le temps que le compteur libère son <c>:FETCh?</c>
        /// en cours et que la boucle de mesure sorte.
        /// </summary>
        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(StopperMesureCommand))]
        private bool _arretEnCours;

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(RelancerMesureCommand))]
        private bool _derniereMesureDisponible;

        /// <summary>Nombre d'appareils détectés sur le bus GPIB (mis à jour après chaque scan).</summary>
        [ObservableProperty] private int _nbAppareilsDetectes;

        /// <summary>Résumé du scan : "3 détecté(s), 2 reconnu(s)" — affiché sur la carte Diagnostic.</summary>
        [ObservableProperty] private string _resumeScanGpib = "Scan en cours...";

        /// <summary>Vrai tant que le scan initial n'est pas terminé (cache le texte détaillé).</summary>
        [ObservableProperty] private bool _scanInitialEnCours = true;

        private double? _derniereFNominale;

        public AccueilViewModel()
        {
            _orchestrator = new MesureOrchestrator(_ieeeDriver, _excelService);

            // Se tient à jour si l'administrateur change le rubidium actif
            EtatApplication.RubidiumActifChange += (_, _) =>
            {
                RubidiumActifTexte = EtatApplication.RubidiumActifTexte;
            };

            // Se met à jour si un scan est relancé depuis Diagnostic GPIB (ajout d'un appareil).
            EtatApplication.AppareilsDetectesChange += (_, _) => RafraichirResumeScan();

            // Scan GPIB initial en arrière-plan — ne bloque pas l'ouverture de la fenêtre.
            _ = Task.Run(ScannerInitialAsync);
        }

        /// <summary>
        /// Relance manuellement un scan GPIB depuis la page d'accueil (bouton Rescanner).
        /// Utile si l'utilisateur branche un appareil en cours de session ou si le scan initial
        /// n'a pas trouvé ce qu'il attendait.
        /// </summary>
        [RelayCommand]
        private Task RescannerAsync()
        {
            Log("🔄 Relance du scan GPIB...");
            return ScannerInitialAsync();
        }

        /// <summary>
        /// Scan GPIB lancé automatiquement au démarrage. Résultat : <see cref="NbAppareilsDetectes"/>
        /// + texte résumé dans <see cref="InformationsGenerales"/>. L'utilisateur n'a pas besoin de
        /// scanner manuellement à chaque démarrage.
        /// </summary>
        private async Task ScannerInitialAsync()
        {
            try
            {
                ScanInitialEnCours = true;
                ResumeScanGpib = "Scan en cours...";

                Journal.Info(CategorieLog.Systeme, "SCAN_AUTO_DEBUT",
                    "Scan GPIB automatique au démarrage de Metrologo.");

                // Timeout étendu à 3s : certains appareils (53131A en particulier) mettent du temps
                // à répondre au *IDN? après un POWER-ON récent ou si le bus a été perturbé.
                var resultats = await ScannerGpib.ScannerAsync(gpibBoard: 0, timeoutMs: 3000);

                // Met à jour l'état global des appareils détectés (consommé par Configuration, etc.)
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

            // Avertissement : réponse *IDN? incohérente = possible conflit d'adresse (deux
            // appareils réglés sur la même adresse GPIB).
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

        /// <summary>
        /// Ouvre dans l'Explorateur Windows le dossier où Metrologo stocke les fichiers Excel
        /// générés (<c>%USERPROFILE%\Documents\Metrologo</c>). Crée le dossier s'il n'existe pas.
        /// </summary>
        [RelayCommand]
        private void OuvrirDossierMesures()
        {
            // Priorité au chemin réseau configuré (Admin > Chemins d'accès > Mesures —
            // chemin réseau partagé). Fallback sur le dossier local Bureau\Metrologo
            // si le réseau est vide ou inaccessible (coupure connexion, partage down).
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
                catch { /* réseau down, fallback local */ }
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
                // L'assignation ci-dessus garde souvent la même référence (vm avait reçu
                // notre instance par référence) — CommunityToolkit ne notifie pas dans ce
                // cas et le bandeau « Fiche en cours » ne s'affiche jamais. On force.
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

                // Journal utilisateur FI : entrée détaillée avec les paramètres clés.
                string gateLib = Metrologo.Services.EnTetesMesureHelper.LibelleGate(MesureConfig.GateIndex);
                JournalFIService.Ecrire("CONFIG_VALIDEE",
                    $"{MesureConfig.TypeMesure} · {MesureConfig.NbMesures} mesures · "
                    + $"{nomAppareil} · gate {gateLib} · {MesureConfig.ModeMesure} · {MesureConfig.SourceMesure}");

                // Enchaînement automatique : après validation de la config, on déclenche
                // directement le lancement de mesure (ouvrirConfigAvant=false pour éviter
                // une boucle de réouverture de la fenêtre de config).
                await ExecuterMesureInterneAsync(ouvrirConfigAvant: false);
            }
        }

        /// <summary>
        /// Commande déclenchée par le bouton « Lancer la mesure ». L'utilisateur veut
        /// systématiquement passer par la fenêtre de configuration avant chaque mesure
        /// (même si une config précédente existe) — on délègue donc à OuvrirConfigurationAsync
        /// qui rappellera <see cref="ExecuterMesureInterneAsync"/>(false) après validation.
        /// Pour relancer SANS reconfigurer, utiliser le bouton « Relancer ».
        /// </summary>
        [RelayCommand]
        private Task ExecuterMesureAsync() => ExecuterMesureInterneAsync(ouvrirConfigAvant: true);

        private async Task ExecuterMesureInterneAsync(bool ouvrirConfigAvant)
        {
            if (MesureEnCours) return;

            // 1) Rubidium obligatoire — défini uniquement dans le menu Administration
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

            // 2) Cycle 1 — clic sur « Lancer la mesure » : on ouvre la config avant la mesure,
            //    même si une config valide existait déjà (l'utilisateur reconfigure à chaque
            //    mesure : N° FI différent, paramètres modifiés…). À la validation,
            //    OuvrirConfigurationAsync rappelle ExecuterMesureInterneAsync(false) pour
            //    enchaîner sur le lancement réel. On RETURN ici pour éviter le double-call.
            if (ouvrirConfigAvant || string.IsNullOrWhiteSpace(MesureConfig?.NumFI))
            {
                Log("ℹ Ouverture de la configuration avant lancement.");
                await OuvrirConfigurationAsync();
                return;
            }

            // 2bis) Module d'incertitude obligatoire — sans module sélectionné, les
            // coefficients A/B (et C/D pour tachy) restent à des valeurs par défaut
            // hardcoded, ce qui invalide le calcul d'incertitude global du rapport.
            // On refuse le lancement et on rouvre la fenêtre Configuration pour que
            // l'utilisateur fasse son choix.
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

            // 2ter) En tachymétrie, le module Fréquence auxiliaire (pour ZNCoeffA/B en Hz)
            // est également obligatoire — il caractérise le fréquencemètre support, donc
            // sans lui le rapport est incomplet côté Hz.
            if (EnTetesMesureHelper.EstTachymetre(MesureConfig.TypeMesure)
                && string.IsNullOrWhiteSpace(MesureConfig.NumModuleIncertitudeFreq))
            {
                Log("✖ Mesure tachy impossible : module Fréquence auxiliaire non sélectionné.");
                Journal.Warn(CategorieLog.Mesure, "MESURE_BLOQUEE_MODULE_FREQ",
                    "Tentative de mesure tachy sans module Fréquence auxiliaire.");
                MessageBox.Show(
                    "Pour une mesure tachymétrique, il faut sélectionner DEUX modules :\n"
                    + "  • Module d'incertitude (tachy) — pour les Coeff C/D côté RPM\n"
                    + "  • Module Fréquence (Hz) — pour les Coeff A/B côté Hz (fréquencemètre)\n\n"
                    + "Le second est manquant. Va dans la fenêtre Configuration et choisis "
                    + "le module Fréquence correspondant au compteur utilisé.",
                    "Module Fréquence requis", MessageBoxButton.OK, MessageBoxImage.Warning);
                await OuvrirConfigurationAsync();
                return;
            }

            // 3) Gate — déjà sélectionné dans la fenêtre Configuration (MesureConfig.GateIndex).
            //    Pour les mesures de Stabilité, on ouvre la fenêtre dédiée pour que l'utilisateur
            //    choisisse les gates à balayer (1 ou plusieurs, via cases à cocher + presets).
            if (MesureConfig.TypeMesure == TypeMesure.Stabilite)
            {
                var gateWin = new SelectionGateWindow(MesureConfig) { Owner = Application.Current.MainWindow };
                if (gateWin.ShowDialog() != true) { Log("✖ Mesure annulée (sélection gates)."); return; }
                MesureConfig.GateIndices = gateWin.ViewModel.IndicesGatesResultats;
            }
            else if (MesureConfig.GateIndices.Count > 1)
            {
                // Filet de sécurité : si on lance une mesure non-Stab juste après une Stab,
                // les multiples gates sélectionnées par la Stab sont encore en mémoire.
                // On force la liste à un seul élément (la 1ère gate) — sinon la mesure
                // Fréquence boucle sur toutes les gates de la Stab précédente.
                MesureConfig.GateIndex = MesureConfig.GateIndices[0]; // setter remet la liste à 1 élément
            }
            Log($"⏱ Gates à balayer : {string.Join(", ", MesureConfig.GateIndices)}");

            // 4) La fréquence nominale est déjà saisie dans ConfigurationWindow (bloc Indirect),
            //    conforme au Delphi d'origine (pas de dialog pré-mesure pour FNominale).
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

            // Choix : conserver les anciennes mesures (= ajouter une feuille au fichier
            // existant) ou écraser (= repartir d'un fichier neuf à partir du template).
            //   Oui = écraser (fichier supprimé)
            //   Non = conserver (= comportement par défaut, mesures ajoutées à la suite)
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
                // OPTION A v2 : si un classeur est ouvert dans Excel COM, on le réinitialise IN PLACE
                // (suppression des feuilles freq*/stab*/etc. + nettoyage Récap.) SANS jamais fermer
                // le Workbook. Évite la fenêtre grise SDI qui apparaissait avec l'ancien flux
                // FermerClasseurActif + File.Delete (~1-3s à 0 Workbook = shell garantie).
                //
                // Si aucun classeur n'est ouvert (cas rare : l'utilisateur a fermé Excel manuellement
                // après la mesure précédente), on retombe sur le flux historique fermer + delete.
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
                    // Pas de classeur ouvert — flux historique (rien à éviter, pas de shell ici).
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
            // Vérifie si un Excel externe (autre que notre instance COM cachée) est ouvert :
            // il peut tenir verrouillé le fichier .xlsx de la mesure et faire échouer ClosedXML.
            // On distingue les reliquats COM sans fenêtre (fantômes) des vrais classeurs ouverts.
            var (_, excelsFantomes) = ExcelInteropHost.Instance.ListerExcelsExternesClasses();

            // 1. Fantômes (aucune fenêtre = reliquat de pilotage COM, jamais un document
            //    utilisateur) : on les ferme en silence pour libérer un éventuel verrou.
            if (excelsFantomes.Count > 0)
            {
                int nbFantomes = ExcelInteropHost.Instance.FermerExcels(excelsFantomes);
                Journal.Info(CategorieLog.Excel, "EXCELS_FANTOMES_NETTOYES",
                    $"{nbFantomes} processus EXCEL.EXE résiduel(s) (sans fenêtre) fermé(s) "
                  + "automatiquement avant la mesure.");
            }

            // 2. Classeurs RÉELLEMENT ouverts par l'utilisateur : on ne tue PLUS l'instance
            //    Excel (cela fermait aussi ses fichiers persos non sauvegardés). On ferme
            //    UNIQUEMENT les classeurs de mesure — qui portent toujours le nom freq/stab —
            //    par leur nom, via la Running Object Table. Tout autre fichier ouvert par
            //    l'utilisateur reste intact. Ça suffit à libérer le verrou sur le rapport de la FI.
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

            // Affiche la fenêtre flottante « ARRÊTER LA MESURE » sur un THREAD UI
            // SÉPARÉ (STA + Dispatcher.Run). Garantit que le clic/raccourci est traité
            // instantanément même quand le thread principal est saturé par les Interop
            // Excel COM. Le callback onStop marshale vers le Dispatcher principal pour
            // exécuter la séquence d'arrêt (annulation CTS, abort GPIB, fermeture Excel).
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
                        MessageBox.Show(
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
                    MessageBox.Show(result.Erreur ?? "Erreur inconnue.",
                        "Mesure interrompue", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                Log($"✖ Erreur inattendue : {ex.Message}");
                Journal.Erreur(CategorieLog.Mesure, "MESURE_EXCEPTION", ex.Message, new { ex.StackTrace });
                JournalFIService.Ecrire("MESURE_ECHEC", $"Exception : {ex.Message}");
                MessageBox.Show(ex.Message, "Erreur inattendue",
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
        ///   <item>TachyContact → page minimale résolution (tr/min) uniquement.</item>
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

            if (config.TypeMesure == TypeMesure.TachyContact
                || config.TypeMesure == TypeMesure.TachyOptique)
            {
                await SaisiePostMesureTachyAsync(config);
                return;
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
