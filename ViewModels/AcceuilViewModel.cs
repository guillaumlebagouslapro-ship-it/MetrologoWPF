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
        // Driver IEEE bas niveau : VISA réel via NI-VISA.
        // Pour revenir à la simulation sans matériel, remplacer par `new SimulationIeeeDriver()`.
        private readonly IIeeeDriver _ieeeDriver = new VisaIeeeDriver(gpibBoard: 0);
        // Type concret conservé (au lieu de IExcelService) pour accéder à FallbackTimestampUtilise.
        private readonly ExcelService _excelService = new ExcelService();
        private readonly MesureOrchestrator _orchestrator;

        private CancellationTokenSource? _cts;

        [ObservableProperty] private bool _estSurBaie = true;
        [ObservableProperty] private string _informationsGenerales = "Prêt. En attente d'exécution...";
        [ObservableProperty] private string _rubidiumActifTexte = EtatApplication.RubidiumActifTexte;
        [ObservableProperty] private Mesure _mesureConfig = new Mesure();

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(RelancerMesureCommand))]
        [NotifyCanExecuteChangedFor(nameof(StopperMesureCommand))]
        private bool _mesureEnCours;

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
                            ModeleReconnu = modeleReconnu
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
        }

        // -------- Commandes --------

        /// <summary>
        /// Ouvre dans l'Explorateur Windows le dossier où Metrologo stocke les fichiers Excel
        /// générés (<c>%USERPROFILE%\Documents\Metrologo</c>). Crée le dossier s'il n'existe pas.
        /// </summary>
        [RelayCommand]
        private void OuvrirDossierMesures()
        {
            string dossier = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "Metrologo");

            try
            {
                System.IO.Directory.CreateDirectory(dossier);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = dossier,
                    UseShellExecute = true
                });
                Journal.Info(CategorieLog.Systeme, "OUVERTURE_DOSSIER_MESURES",
                    $"Ouverture du dossier des mesures : {dossier}");
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
        private void OuvrirConfiguration()
        {
            if (MesureConfig == null) MesureConfig = new Mesure();

            var vm = new ConfigurationViewModel { MesureConfig = MesureConfig };
            vm.EstSurBaie = EstSurBaie;

            var win = new ConfigurationWindow(vm) { Owner = Application.Current.MainWindow };
            if (win.ShowDialog() == true)
            {
                MesureConfig = vm.MesureConfig;
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
            }
        }

        [RelayCommand]
        private async Task ExecuterMesureAsync()
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

            // 2) Configuration (FI obligatoire)
            if (string.IsNullOrWhiteSpace(MesureConfig?.NumFI))
            {
                Log("ℹ Configuration requise avant de lancer une mesure.");
                OuvrirConfiguration();
                if (MesureConfig == null || string.IsNullOrWhiteSpace(MesureConfig.NumFI))
                {
                    Log("✖ Mesure annulée (configuration incomplète).");
                    return;
                }
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
            await LancerMesureAsync(MesureConfig, rubi, _derniereFNominale,
                preambule: "🔁 Relance (mêmes paramètres)");
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
            _cts = new CancellationTokenSource();
            MesureEnCours = true;

            string nomAppareilLog = ResolvedNomAppareil(config);
            Log("═══════════════════════════════════════════");
            Log($"{preambule} : {config.NbMesures} mesures sur {nomAppareilLog}");
            Log($"   FI {config.NumFI} · Rubidium : {rubi.Designation} · "
              + (rubi.AvecGPS ? "GPS" : "Allouis"));

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

            var progress = new Progress<ProgressionMesure>(p =>
            {
                if (p.DerniereValeur.HasValue)
                    Log($"   {p.Message} : {p.DerniereValeur.Value:F6} Hz");
                else
                    Log($"… {p.Message}");
            });

            try
            {
                var result = await _orchestrator.ExecuterAsync(
                    config, rubi, fNominale, progress, _cts.Token);

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

                    // Saisie post-mesure : fréquence lue + incertitudes — uniquement pour
                    // Fréquence + Source=Fréquencemètre (conforme Delphi : skip si Générateur).
                    await SaisiePostMesureAsync(config);
                }
                else
                {
                    Log($"✖ Échec : {result.Erreur}");
                    Journal.Erreur(CategorieLog.Mesure, "MESURE_ECHEC", result.Erreur ?? "Échec inconnu.");
                    MessageBox.Show(result.Erreur ?? "Erreur inconnue.",
                        "Mesure interrompue", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                Log($"✖ Erreur inattendue : {ex.Message}");
                Journal.Erreur(CategorieLog.Mesure, "MESURE_EXCEPTION", ex.Message, new { ex.StackTrace });
                MessageBox.Show(ex.Message, "Erreur inattendue",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                MesureEnCours = false;
                _cts?.Dispose();
                _cts = null;
            }
        }

        /// <summary>
        /// Flux post-mesure conforme au Delphi (F_Main.pas:2283) : pour une mesure de Fréquence
        /// avec le fréquencemètre comme source, demande à l'utilisateur la fréquence lue sur
        /// l'afficheur puis les incertitudes (résolution + autres). Les valeurs sont écrites
        /// dans les zones nommées du classeur Excel via Interop.
        /// Skip pour Générateur (incertitudes à 0) et pour tous les autres types de mesure.
        /// </summary>
        private async Task SaisiePostMesureAsync(Mesure config)
        {
            if (config.TypeMesure != TypeMesure.Frequence) return;
            if (config.SourceMesure != SourceMesure.Frequencemetre) return;

            // 1) Fréquence lue sur l'afficheur
            var saisieLue = new SaisieValFreqViewModel(
                config.FNominale,
                titre: "Fréquence lue",
                sousTitre: "Saisissez la valeur affichée par le fréquencemètre",
                libelle: "FRÉQUENCE LUE (HZ)");

            var winLue = new SaisieValFreqWindow(saisieLue) { Owner = Application.Current.MainWindow };
            if (winLue.ShowDialog() != true)
            {
                Log("ℹ Saisie post-mesure ignorée (fréquence lue).");
                return;
            }

            // 2) Incertitude de résolution + autres incertitudes
            var paramsVm = new ParamsIncertViewModel(config.Resolution, config.IncertSupp);
            var winIncert = new ParamsIncertWindow(paramsVm) { Owner = Application.Current.MainWindow };
            if (winIncert.ShowDialog() != true)
            {
                Log("ℹ Saisie post-mesure ignorée (incertitudes).");
                return;
            }

            config.Resolution = paramsVm.Resolution;
            config.IncertSupp = paramsVm.IncertSupp;

            Log($"📝 Post-mesure : FLue={saisieLue.ValeurLue:F6} Hz · Résolution={paramsVm.Resolution:E3} · Autres={paramsVm.IncertSupp:E3}");

            // Écrit via Interop dans le classeur actif (zones nommées du template).
            await ExcelInteropHost.Instance.EcrireZoneNommeeAsync("ZNFLue", saisieLue.ValeurLue);
            await ExcelInteropHost.Instance.EcrireZoneNommeeAsync("ZNIncertResol", paramsVm.Resolution);
            await ExcelInteropHost.Instance.EcrireZoneNommeeAsync("ZNIncertSup", paramsVm.IncertSupp);
            await ExcelInteropHost.Instance.SauvegarderAsync();
        }

        [RelayCommand(CanExecute = nameof(PeutStopper))]
        private void StopperMesure()
        {
            if (_cts != null && !_cts.IsCancellationRequested)
            {
                _cts.Cancel();
                Log("⏹ Arrêt demandé — en cours…");
                Journal.Warn(CategorieLog.Mesure, "MESURE_STOP", "Arrêt demandé par l'utilisateur.");
            }
        }

        private bool PeutStopper() => MesureEnCours;

        // -------- Utilitaires --------

        private void Log(string message)
        {
            InformationsGenerales += $"\n[{DateTime.Now:HH:mm:ss}] {message}";
        }
    }
}
