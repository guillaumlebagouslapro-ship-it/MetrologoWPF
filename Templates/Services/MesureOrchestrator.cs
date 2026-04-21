using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Metrologo.Models;

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
        private readonly IIeeeService _ieee;
        private readonly IExcelService _excel;

        public MesureOrchestrator(IIeeeService ieee, IExcelService excel)
        {
            _ieee = ieee;
            _excel = excel;
        }

        public async Task<ResultatMesure> ExecuterAsync(
            Mesure config,
            Rubidium rubidium,
            double? fNominaleOuReference,
            IProgress<ProgressionMesure>? progress = null,
            CancellationToken ct = default)
        {
            var result = new ResultatMesure();

            try
            {
                progress?.Report(new ProgressionMesure
                {
                    Message = "Initialisation du bus IEEE...",
                    EtapeActuelle = 0,
                    EtapesTotales = config.NbMesures + 2
                });

                if (!await _ieee.InitialiserAsync())
                {
                    result.Erreur = "Impossible d'initialiser le bus IEEE.";
                    return result;
                }

                // Saisie de la fréquence nominale si nécessaire
                if (fNominaleOuReference.HasValue)
                {
                    config.FNominale = fNominaleOuReference.Value;
                    result.FNominale = fNominaleOuReference.Value;
                }

                // Préparation du rapport Excel
                progress?.Report(new ProgressionMesure
                {
                    Message = $"Préparation du rapport Excel pour FI {config.NumFI}...",
                    EtapeActuelle = 1,
                    EtapesTotales = config.NbMesures + 2
                });
                await _excel.InitialiserRapportAsync(config.NumFI, config, rubidium);

                // Boucle de mesures
                var valeurs = new List<double>();
                for (int i = 0; i < config.NbMesures; i++)
                {
                    ct.ThrowIfCancellationRequested();

                    double val = await _ieee.LireMesureAsync(config, ct);
                    valeurs.Add(val);

                    progress?.Report(new ProgressionMesure
                    {
                        Message = $"Mesure {i + 1}/{config.NbMesures}",
                        EtapeActuelle = i + 2,
                        EtapesTotales = config.NbMesures + 2,
                        DerniereValeur = val
                    });
                }

                result.Valeurs = valeurs;

                // Transfert Excel
                progress?.Report(new ProgressionMesure
                {
                    Message = "Transfert des résultats vers Excel...",
                    EtapeActuelle = config.NbMesures + 2,
                    EtapesTotales = config.NbMesures + 2
                });
                await _excel.AjouterResultatsAsync(valeurs);
                await _excel.SauvegarderEtOuvrirAsync();

                result.Succes = true;
            }
            catch (OperationCanceledException)
            {
                result.Erreur = "Mesure annulée par l'utilisateur.";
            }
            catch (Exception ex)
            {
                result.Erreur = ex.Message;
            }
            finally
            {
                _excel.FermerExcel();
            }

            return result;
        }
    }
}
