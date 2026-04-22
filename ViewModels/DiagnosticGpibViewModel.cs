using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Ieee;
using Metrologo.Services.Journal;
using Metrologo.Views;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// ViewModel du diagnostic GPIB : scanne le bus via NI-VISA et affiche les appareils détectés.
    /// </summary>
    public partial class DiagnosticGpibViewModel : ObservableObject
    {
        [ObservableProperty] private string _statut = "Prêt. Cliquez sur « Scanner le bus » pour détecter les appareils.";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(PeutScanner))]
        private bool _scanEnCours;

        [ObservableProperty] private int _nbAppareilsDetectes;
        [ObservableProperty] private int _nbErreurs;
        [ObservableProperty] private string _messageErreur = string.Empty;
        [ObservableProperty] private bool _aErreur;
        [ObservableProperty] private string _ressourcesVisa = string.Empty;
        [ObservableProperty] private bool _ressourcesVisaAffiche;

        public bool PeutScanner => !ScanEnCours;
        public ObservableCollection<ResultatScanGpib> Resultats { get; } = new();

        public DiagnosticGpibViewModel()
        {
            // Réagit aux changements du catalogue : si un modèle est ajouté, les lignes du scan
            // en cours sont réactualisées pour refléter la reconnaissance.
            CatalogueAppareilsService.Instance.CatalogueChange += (_, _) => RafraichirReconnaissance();
        }

        private void RafraichirReconnaissance()
        {
            foreach (var r in Resultats.Where(x => x.Repond))
                r.ModeleCatalogue = CatalogueAppareilsService.Instance.TrouverParIdn(r.Fabricant, r.Modele);
        }

        [RelayCommand]
        private async Task ScannerAsync()
        {
            if (ScanEnCours) return;

            Resultats.Clear();
            NbAppareilsDetectes = 0;
            NbErreurs = 0;
            AErreur = false;
            MessageErreur = string.Empty;
            ScanEnCours = true;
            Statut = "Scan en cours — détection via NI-VISA puis interrogation *IDN?...";

            JournalLog.Info(CategorieLog.Administration, "SCAN_GPIB_DEBUT",
                "Début du scan du bus GPIB.");

            try
            {
                var progress = new Progress<ResultatScanGpib>(r =>
                {
                    // Matching catalogue : l'appareil est-il déjà enregistré ?
                    if (r.Repond)
                        r.ModeleCatalogue = CatalogueAppareilsService.Instance
                            .TrouverParIdn(r.Fabricant, r.Modele);

                    Resultats.Add(r);
                    if (r.Repond) NbAppareilsDetectes++;
                    if (!string.IsNullOrEmpty(r.Erreur)) NbErreurs++;
                    Statut = $"Interrogation GPIB{r.Board}::{r.Adresse} — {NbAppareilsDetectes} appareil(s) détecté(s)"
                             + (NbErreurs > 0 ? $", {NbErreurs} erreur(s)" : "");
                });

                await ScannerGpib.ScannerAsync(
                    gpibBoard: 0,
                    timeoutMs: 2000,
                    progress: progress,
                    ct: CancellationToken.None);

                Statut = NbAppareilsDetectes == 0
                    ? NbErreurs > 0
                        ? $"Scan terminé : aucun appareil détecté, {NbErreurs} erreur(s). Voir détails ci-dessous."
                        : "Scan terminé : aucun appareil détecté sur le bus GPIB. Essayez « Lister les ressources VISA »."
                    : $"Scan terminé : {NbAppareilsDetectes} appareil(s) détecté(s).";

                // Met à jour l'état global partagé (visible dans la fenêtre Configuration)
                EtatApplication.AppareilsDetectes.Clear();
                foreach (var r in Resultats.Where(r => r.Repond))
                {
                    EtatApplication.AppareilsDetectes.Add(new AppareilDetecte
                    {
                        Board = r.Board,
                        Adresse = r.Adresse,
                        Ressource = r.Ressource,
                        IdnBrut = r.ReponseIdn,
                        Fabricant = r.Fabricant,
                        Modele = r.Modele,
                        NumeroSerie = r.NumeroSerie,
                        Firmware = r.Firmware,
                        TypeReconnu = AppareilDetecte.DeviquerType(r.Fabricant, r.Modele),
                        ModeleReconnu = r.ModeleCatalogue
                    });
                }
                EtatApplication.NotifierAppareilsDetectesChange();

                JournalLog.Info(CategorieLog.Administration, "SCAN_GPIB_FIN",
                    $"Scan terminé : {NbAppareilsDetectes} appareil(s) détecté(s), {NbErreurs} erreur(s).",
                    new { NbAppareils = NbAppareilsDetectes, NbErreurs });
            }
            catch (Exception ex)
            {
                AErreur = true;
                MessageErreur = ex.Message;
                Statut = "Échec du scan.";
                JournalLog.Erreur(CategorieLog.Administration, "SCAN_GPIB_ERREUR",
                    $"Erreur lors du scan : {ex.Message}",
                    new { ex.GetType().Name });
            }
            finally
            {
                ScanEnCours = false;
            }
        }

        [RelayCommand]
        private async Task ListerRessourcesVisaAsync()
        {
            if (ScanEnCours) return;
            ScanEnCours = true;
            Statut = "Interrogation de NI-VISA en cours...";
            AErreur = false;
            RessourcesVisaAffiche = false;

            try
            {
                var ressources = await ScannerGpib.ListerRessourcesAsync();
                if (ressources.Count == 0)
                {
                    RessourcesVisa = "Aucune ressource VISA trouvée. NI-VISA ne voit aucun appareil "
                                   + "— vérifie le câble GPIB, l'alimentation de l'appareil, et l'adaptateur USB.";
                }
                else
                {
                    RessourcesVisa = string.Join(Environment.NewLine, ressources);
                }
                RessourcesVisaAffiche = true;
                Statut = $"NI-VISA connaît {ressources.Count} ressource(s).";

                JournalLog.Info(CategorieLog.Administration, "LISTE_RESSOURCES_VISA",
                    $"{ressources.Count} ressource(s) VISA.",
                    new { Ressources = ressources });
            }
            catch (Exception ex)
            {
                AErreur = true;
                MessageErreur = $"Impossible d'interroger NI-VISA : {ex.Message}";
                Statut = "Échec de l'interrogation de NI-VISA.";
            }
            finally
            {
                ScanEnCours = false;
            }
        }

        [RelayCommand]
        private void EnregistrerAppareil(ResultatScanGpib? resultat)
        {
            if (resultat == null || !resultat.Repond) return;

            var detecte = new AppareilDetecte
            {
                Board = resultat.Board,
                Adresse = resultat.Adresse,
                Ressource = resultat.Ressource,
                IdnBrut = resultat.ReponseIdn,
                Fabricant = resultat.Fabricant,
                Modele = resultat.Modele,
                NumeroSerie = resultat.NumeroSerie,
                Firmware = resultat.Firmware
            };

            string utilisateur = JournalLog.Utilisateur ?? "inconnu";
            var vm = new EnregistrementAppareilViewModel(detecte, utilisateur);
            var win = new EnregistrementAppareilWindow(vm) { Owner = Application.Current.MainWindow };
            win.ShowDialog();
            // Le CatalogueChange déclenché par AjouterAsync rafraîchit déjà la reconnaissance.
        }

        [RelayCommand]
        private void Fermer()
        {
            foreach (Window w in Application.Current.Windows)
            {
                if (w.DataContext == this) { w.Close(); return; }
            }
        }
    }
}
