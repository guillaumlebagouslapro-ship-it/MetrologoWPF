using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Metrologo.Models;
using Metrologo.Services.Catalogue;
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
                // 0. Réinitialise les sessions GPIB cachées par le driver VISA — évite les
                //    timeouts quand un appareil a été éteint/rallumé entre deux mesures.
                _driver.ReinitialiserSessions();

                // 1. Résolution de l'appareil :
                //    - priorité au catalogue local si mesure.IdModeleCatalogue est défini (ex: Agilent 53131A)
                //    - sinon fallback sur Metrologo.ini (Stanford / Racal / EIP historiques)
                AppareilIEEE appareil;
                if (!string.IsNullOrWhiteSpace(mesure.IdModeleCatalogue))
                {
                    var (appareilCatalogue, erreur) = ResolverDepuisCatalogue(mesure.IdModeleCatalogue);
                    if (appareilCatalogue == null)
                    {
                        result.Erreur = erreur;
                        return result;
                    }
                    appareil = appareilCatalogue;
                }
                else
                {
                    var configApp = EtatApplication.ConfigAppareils;
                    if (configApp == null)
                    {
                        result.Erreur = "Configuration des appareils non chargée. "
                            + "Vérifiez que Metrologo.ini est présent dans le dossier Config.";
                        return result;
                    }
                    appareil = configApp.Par(mesure.Frequencemetre);
                }

                progress?.Report(new ProgressionMesure
                {
                    Message = $"Initialisation du {appareil.Nom} (GPIB {appareil.Adresse})...",
                    EtapeActuelle = 0,
                    EtapesTotales = mesure.NbMesures + 3
                });

                // 2. Initialisation de l'appareil (envoi ChaineInit)
                await appareil.InitialiserAsync(_driver, ct);

                // 3. Configuration : IFC + MUX + ConfEntree + activation SRQ
                //    MUX uniquement disponible pour les appareils legacy (chargé depuis Metrologo.ini).
                var mux = EtatApplication.ConfigAppareils?.Mux;
                await appareil.ConfigurerAsync(_driver, mesure, mux, commandesMux: null, ct);

                // 3 bis. Rejoue les commandes SCPI des réglages dynamiques choisis dans Configuration
                //        (Impédance, Couplage, Filtre, Trigger, Mode…). Sans ça le *RST de la ChaineInit
                //        efface tous les réglages que l'utilisateur a validés avant de lancer la mesure.
                //        Chaque commande est wrappée : si une timeout, on logge et on continue les
                //        suivantes (plutôt que de planter toute la mesure). Un petit délai entre
                //        commandes évite de saturer le handshake GPIB sur les appareils lents.
                if (mesure.CommandesScpiReglages != null && mesure.CommandesScpiReglages.Count > 0)
                {
                    foreach (var cmd in mesure.CommandesScpiReglages)
                    {
                        if (string.IsNullOrWhiteSpace(cmd)) continue;

                        try
                        {
                            await _driver.EcrireAsync(appareil.Adresse, cmd, appareil.WriteTerm, ct);
                            JournalLog.Info(CategorieLog.Mesure, "SCPI_REJEU",
                                $"GPIB0::{appareil.Adresse} ← {cmd} (réapplication post-RST)");
                        }
                        catch (Ivi.Visa.IOTimeoutException)
                        {
                            JournalLog.Warn(CategorieLog.Mesure, "SCPI_REJEU_TIMEOUT",
                                $"Timeout sur l'envoi de « {cmd } » à GPIB0::{appareil.Adresse} — "
                                + "commande ignorée, la mesure continue.");
                        }

                        // Laisse l'appareil digérer (50 ms est un compromis testé pour 53131A / Stanford SR620).
                        await Task.Delay(50, ct);
                    }
                }

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

                // 6. Boucle de mesures — on mémorise le timestamp RÉEL juste après chaque lecture
                //    GPIB pour qu'Excel affiche les vrais horodatages (et non une heure + i secondes
                //    factice calculée à la fin).
                var valeurs = new List<double>();
                var horodatages = new List<DateTime>();
                for (int i = 0; i < mesure.NbMesures; i++)
                {
                    ct.ThrowIfCancellationRequested();

                    double val = await appareil.MesurerAsync(_driver, mesure, ct);
                    valeurs.Add(val);
                    horodatages.Add(DateTime.Now);

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
                await _excel.AjouterResultatsAsync(valeurs, horodatages);
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

        /// <summary>
        /// Retrouve le modèle catalogue + l'adresse GPIB détectée sur le bus, et construit
        /// un <see cref="AppareilIEEE"/> via <see cref="CatalogueAdapter"/>. Retourne un message
        /// d'erreur explicite si le modèle est introuvable ou si l'appareil n'est pas détecté.
        /// </summary>
        private static (AppareilIEEE?, string?) ResolverDepuisCatalogue(string idModele)
        {
            var modele = CatalogueAppareilsService.Instance.Modeles
                .FirstOrDefault(m => m.Id == idModele);
            if (modele == null)
            {
                return (null, $"Le modèle catalogue « {idModele} » est introuvable. "
                    + "Le catalogue a peut-être été modifié — rouvrez la Configuration.");
            }

            var detecte = EtatApplication.AppareilsDetectes
                .FirstOrDefault(d => modele.Correspond(d.Fabricant, d.Modele));
            if (detecte == null)
            {
                return (null, $"Le modèle « {modele.Nom } » n'est pas détecté sur le bus GPIB. "
                    + "Lancez un scan depuis Diagnostic GPIB pour le retrouver.");
            }

            return (CatalogueAdapter.VersAppareilIEEE(modele, detecte.Adresse), null);
        }
    }
}
