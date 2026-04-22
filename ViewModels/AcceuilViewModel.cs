using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Services.Ieee;
using Metrologo.Services.Journal;
using Metrologo.Views;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace Metrologo.ViewModels
{
    public partial class AccueilViewModel : ObservableObject
    {
        // Driver IEEE bas niveau — pour l'instant la simulation, à remplacer par VisaIeeeDriver
        // quand le driver NI-VISA sera branché.
        private readonly IIeeeDriver _ieeeDriver = new SimulationIeeeDriver();
        private readonly IExcelService _excelService = new ExcelService();
        private readonly MesureOrchestrator _orchestrator;

        private CancellationTokenSource? _cts;

        [ObservableProperty] private bool _estSurBaie = true;
        [ObservableProperty] private string _informationsGenerales = "Prêt. En attente d'exécution...";
        [ObservableProperty] private string _rubidiumActifTexte = EtatApplication.RubidiumActifTexte;
        [ObservableProperty] private Mesure _mesureConfig = new Mesure();

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(RelancerMesureCommand))]
        private bool _mesureEnCours;

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(RelancerMesureCommand))]
        private bool _derniereMesureDisponible;

        private double? _derniereFNominale;

        public AccueilViewModel()
        {
            _orchestrator = new MesureOrchestrator(_ieeeDriver, _excelService);

            // Se tient à jour si l'administrateur change le rubidium actif
            EtatApplication.RubidiumActifChange += (_, _) =>
            {
                RubidiumActifTexte = EtatApplication.RubidiumActifTexte;
            };
        }

        // -------- Commandes --------

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
                Log($"⚙ Configuration : FI {MesureConfig.NumFI} · {MesureConfig.TypeMesure} · "
                  + $"{MesureConfig.Frequencemetre} · {MesureConfig.NbMesures} mesures · "
                  + $"Mode {MesureConfig.ModeMesure}");

                Journal.Info(CategorieLog.Configuration, "CONFIG_VALIDEE",
                    $"FI {MesureConfig.NumFI} · {MesureConfig.TypeMesure} · {MesureConfig.Frequencemetre}",
                    new
                    {
                        numFI = MesureConfig.NumFI,
                        type = MesureConfig.TypeMesure.ToString(),
                        appareil = MesureConfig.Frequencemetre.ToString(),
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
            //    Pour les mesures de Stabilité avec procédure auto (sentinelles -1 / -2 / -3),
            //    on ouvre quand même le dialogue car la Configuration ne gère pas encore ces procédures.
            if (MesureConfig.TypeMesure == TypeMesure.Stabilite)
            {
                var gateWin = new SelectionGateWindow(MesureConfig) { Owner = Application.Current.MainWindow };
                if (gateWin.ShowDialog() != true) { Log("✖ Mesure annulée (procédure)."); return; }
                MesureConfig.GateIndex = gateWin.ViewModel.IndexGateResultat;
            }
            Log($"⏱ Gate : index {MesureConfig.GateIndex}");

            // 4) Saisie fréquence nominale si mode Indirect ou source Générateur
            double? fNominale = null;
            bool besoinNominale = MesureConfig.ModeMesure == ModeMesure.Indirect
                               || MesureConfig.SourceMesure == SourceMesure.Generateur;

            if (besoinNominale)
            {
                var saisieVm = new SaisieValFreqViewModel(
                    MesureConfig.FNominale,
                    titre: MesureConfig.SourceMesure == SourceMesure.Generateur
                        ? "Fréquence du générateur"
                        : "Fréquence nominale",
                    sousTitre: "Saisissez la valeur de référence en Hertz",
                    libelle: "VALEUR (HZ)");

                var saisieWin = new SaisieValFreqWindow(saisieVm) { Owner = Application.Current.MainWindow };
                if (saisieWin.ShowDialog() != true) { Log("✖ Mesure annulée (saisie fréquence)."); return; }

                fNominale = saisieVm.ValeurLue;
                Log($"📝 Fréquence nominale : {fNominale:N3} Hz");
            }

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

        private async Task LancerMesureAsync(Mesure config, Rubidium rubi, double? fNominale, string preambule)
        {
            _cts = new CancellationTokenSource();
            MesureEnCours = true;

            Log("═══════════════════════════════════════════");
            Log($"{preambule} : {config.NbMesures} mesures sur {config.Frequencemetre}");
            Log($"   FI {config.NumFI} · Rubidium : {rubi.Designation} · "
              + (rubi.AvecGPS ? "GPS" : "Allouis"));

            Journal.Info(CategorieLog.Mesure, "MESURE_DEBUT",
                $"{preambule} : {config.NbMesures} mesures sur {config.Frequencemetre} pour FI {config.NumFI}",
                new
                {
                    numFI = config.NumFI,
                    type = config.TypeMesure.ToString(),
                    appareil = config.Frequencemetre.ToString(),
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
                    Log($"✅ Rapport Excel ouvert.");

                    _derniereFNominale = fNominale;
                    DerniereMesureDisponible = true;

                    Journal.Info(CategorieLog.Mesure, "MESURE_FIN",
                        $"Mesure terminée : moyenne {result.Moyenne:F6} Hz, σ {result.EcartType:E3} Hz",
                        new { result.Moyenne, result.EcartType, nbValeurs = result.Valeurs.Count });
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

        [RelayCommand]
        private void StopperMesure()
        {
            if (_cts != null && !_cts.IsCancellationRequested)
            {
                _cts.Cancel();
                Log("⏹ Arrêt demandé — en cours…");
                Journal.Warn(CategorieLog.Mesure, "MESURE_STOP", "Arrêt demandé par l'utilisateur.");
            }
            else
            {
                Log("ℹ Aucune mesure en cours.");
            }
        }

        // -------- Utilitaires --------

        private void Log(string message)
        {
            InformationsGenerales += $"\n[{DateTime.Now:HH:mm:ss}] {message}";
        }
    }
}
