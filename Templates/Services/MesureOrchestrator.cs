using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Metrologo.Models;
using Metrologo.Services.Ieee;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services
{
    public class ResultatMesure
    {
        public bool Succes { get; set; }
        public List<double> Valeurs { get; set; } = new();
        public double? FNominale { get; set; }
        public string? Erreur { get; set; }
        public string? CheminExcel { get; set; }

        public double Moyenne => Valeurs.Count == 0 ? 0 : Moy(Valeurs);
        public double EcartType => Valeurs.Count < 2 ? 0 : Sigma(Valeurs);

        private static double Moy(List<double> v)
        {
            double s = 0; foreach (var x in v) s += x; return s / v.Count;
        }

        private static double Sigma(List<double> v)
        {
            var m = Moy(v);
            double s = 0; foreach (var x in v) s += (x - m) * (x - m);
            return Math.Sqrt(s / (v.Count - 1));
        }
    }

    public class ProgressionMesure
    {
        public string Message { get; set; } = string.Empty;
        public int EtapeActuelle { get; set; }
        public int EtapesTotales { get; set; }
        public double? DerniereValeur { get; set; }
    }

    public class MesureOrchestrator
    {
        private readonly IIeeeDriver _driver;
        private readonly IExcelService _excel;

        public MesureOrchestrator(IIeeeDriver driver, IExcelService excel)
        {
            _driver = driver;
            _excel = excel;
        }

        public async Task<ResultatMesure> ExecuterAsync(
            Mesure mesure,
            Rubidium rubidium,
            double? fNominaleOuReference,
            IProgress<ProgressionMesure>? progress = null,
            CancellationToken ct = default)
        {
            var result = new ResultatMesure();

            try
            {
                // 1. Résolution de l'appareil depuis la config chargée depuis Metrologo.ini
                var configApp = EtatApplication.ConfigAppareils;
                if (configApp == null)
                {
                    result.Erreur = "Configuration des appareils non chargée. "
                        + "Vérifiez que Metrologo.ini est présent dans le dossier Config.";
                    return result;
                }
                var appareil = configApp.Par(mesure.Frequencemetre);

                progress?.Report(new ProgressionMesure
                {
                    Message = $"Initialisation du {appareil.Nom} (GPIB {appareil.Adresse})...",
                    EtapeActuelle = 0,
                    EtapesTotales = mesure.NbMesures + 3
                });

                // 2. Initialisation de l'appareil (envoi ChaineInit)
                await appareil.InitialiserAsync(_driver, ct);

                // 3. Configuration : IFC + MUX + ConfEntree + activation SRQ
                await appareil.ConfigurerAsync(_driver, mesure, configApp.Mux, commandesMux: null, ct);

                // Saisie de la fréquence nominale si nécessaire
                if (fNominaleOuReference.HasValue)
                {
                    mesure.FNominale = fNominaleOuReference.Value;
                    result.FNominale = fNominaleOuReference.Value;
                }

                // 4. Préparation du rapport Excel
                progress?.Report(new ProgressionMesure
                {
                    Message = $"Préparation du rapport Excel pour FI {mesure.NumFI}...",
                    EtapeActuelle = 1,
                    EtapesTotales = mesure.NbMesures + 3
                });
                await _excel.InitialiserRapportAsync(mesure.NumFI, mesure, rubidium);

                // 5. Programmation de la gate sur l'appareil (hors mesure d'intervalle, cf. F_Main:1211)
                progress?.Report(new ProgressionMesure
                {
                    Message = $"Programmation de la gate (index {mesure.GateIndex})...",
                    EtapeActuelle = 2,
                    EtapesTotales = mesure.NbMesures + 3
                });
                await appareil.AppliquerGateAsync(_driver, mesure.GateIndex, mesure.TypeMesure, ct);

                // 6. Boucle de mesures
                var valeurs = new List<double>();
                for (int i = 0; i < mesure.NbMesures; i++)
                {
                    ct.ThrowIfCancellationRequested();

                    double val = await appareil.MesurerAsync(_driver, mesure, ct);
                    valeurs.Add(val);

                    progress?.Report(new ProgressionMesure
                    {
                        Message = $"Mesure {i + 1}/{mesure.NbMesures}",
                        EtapeActuelle = i + 3,
                        EtapesTotales = mesure.NbMesures + 3,
                        DerniereValeur = val
                    });
                }

                result.Valeurs = valeurs;

                // 7. Désactivation SRQ — cf. F_Main:1263 (correctif blocage Racal 10ms↔20ms)
                await appareil.DesactiverSrqAsync(_driver, ct);

                // 8. Mise à jour de la feuille Récap. — portage des MajRecap* du Delphi (F_Main.pas:1276)
                //    Une nouvelle ligne est insérée avec des formules cross-sheet vers la feuille de mesure.
                if (mesure.TypeMesure == TypeMesure.Frequence
                    || mesure.TypeMesure == TypeMesure.FreqAvantInterv
                    || mesure.TypeMesure == TypeMesure.FreqFinale)
                {
                    await _excel.MettreAJourRecapFreqAsync(mesure);
                }
                else if (mesure.TypeMesure == TypeMesure.Stabilite)
                {
                    await _excel.MettreAJourRecapStabAsync(mesure);
                }

                // 9. Transfert Excel
                progress?.Report(new ProgressionMesure
                {
                    Message = "Transfert des résultats vers Excel...",
                    EtapeActuelle = mesure.NbMesures + 3,
                    EtapesTotales = mesure.NbMesures + 3
                });
                await _excel.AjouterResultatsAsync(valeurs);
                await _excel.SauvegarderEtOuvrirAsync();

                JournalLog.Info(CategorieLog.Mesure, "Execute",
                    $"Mesure terminée : {valeurs.Count} valeurs sur {appareil.Nom}.",
                    new { appareil.Nom, appareil.Adresse, GateIndex = mesure.GateIndex, NbMesures = valeurs.Count });

                result.Succes = true;
            }
            catch (OperationCanceledException)
            {
                result.Erreur = "Mesure annulée par l'utilisateur.";
                JournalLog.Warn(CategorieLog.Mesure, "Execute", result.Erreur);
            }
            catch (Exception ex)
            {
                result.Erreur = ex.Message;
                JournalLog.Erreur(CategorieLog.Mesure, "Execute",
                    $"Échec de la mesure : {ex.Message}", new { ex.GetType().Name, ex.StackTrace });
            }
            finally
            {
                _excel.FermerExcel();
            }

            return result;
        }
    }
}
