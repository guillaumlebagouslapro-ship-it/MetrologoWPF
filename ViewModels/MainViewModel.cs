using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Views;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;

namespace Metrologo.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        // Déclaration des services manquants
        private readonly IExcelService _excelService = new ExcelService();
        private readonly IIeeeService _ieeeService = new SimulationIeeeService();
        // Ajoutez ceci au début de la classe MainViewModel
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(EstAdmin))] // Prévient l'interface que les droits changent
        private Utilisateur? _utilisateurConnecte;

        // Cette propriété calculée permet de cacher/afficher des boutons dans le XAML
        public bool EstAdmin => UtilisateurConnecte?.Role == RoleUtilisateur.Administrateur;


        [ObservableProperty]
        private string _titreApplication = "Metrologo v2026 - Migration Delphi";

        [ObservableProperty]
        private string _informationsGenerales = "Prêt. En attente d'exécution...";

        [ObservableProperty]
        private string _rubidiumActifTexte = "Rubidium : Non sélectionné";

        [ObservableProperty]
        private Mesure _mesureConfig = new Mesure(); // Initialisé par défaut pour éviter le null

        [RelayCommand]
        private void OuvrirConfiguration()
        {
            if (MesureConfig == null) MesureConfig = new Mesure();

            var configVM = new ConfigurationViewModel { MesureConfig = this.MesureConfig };
            var win = new ConfigurationWindow(configVM)
            {
                Owner = Application.Current.MainWindow
            };

            if (win.ShowDialog() == true)
            {
                this.MesureConfig = configVM.MesureConfig;
                AjouterInformation($"✅ Configuration validée : FI {MesureConfig.NumFI}, {MesureConfig.NbMesures} mesures.");
            }
        }

        [RelayCommand]
        private async Task ExecuterMesure()
        {
            // 1. On ouvre la fenêtre de configuration AVANT de faire quoi que ce soit
            OuvrirConfiguration();

            // 2. Si l'utilisateur clique sur "Annuler" ou ne met pas de N° FI, on arrête tout
            if (MesureConfig == null || string.IsNullOrWhiteSpace(MesureConfig.NumFI))
            {
                AjouterInformation("ℹ️ Mesure annulée : Configuration incomplète ou abandonnée.");
                return;
            }

            // 3. Si c'est validé, on lance le vrai processus !
            try
            {
                AjouterInformation($"▶ Lancement Excel pour : {MesureConfig.NumFI}");
                await _excelService.InitialiserRapportAsync(MesureConfig.NumFI, MesureConfig);

                AjouterInformation("📡 Connexion aux appareils IEEE...");
                List<double> mesuresRecuperees = new List<double>();

                // Boucle de mesure
                for (int i = 0; i < MesureConfig.NbMesures; i++)
                {
                    double val = await _ieeeService.LireMesureAsync(new AppareilIEEE());
                    mesuresRecuperees.Add(val);
                    AjouterInformation($"   Mesure {i + 1}/{MesureConfig.NbMesures} : {val:F3} Hz");
                }

                AjouterInformation("📝 Transfert vers Excel...");
                await _excelService.AjouterResultatsAsync(mesuresRecuperees);

                await _excelService.SauvegarderEtOuvrirAsync();
                AjouterInformation("✅ Rapport sauvegardé et ouvert avec succès.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur critique : {ex.Message}");
            }
            finally
            {
                _excelService.FermerExcel();
            }
        }

        [RelayCommand]
        private void StopperMesure()
        {
            AjouterInformation("⏹ Arrêt demandé par l'utilisateur.");
        }

        [RelayCommand]
        private void Quitter()
        {
            Application.Current.Shutdown();
        }

        private void AjouterInformation(string message)
        {
            InformationsGenerales += $"\n[{DateTime.Now:HH:mm:ss}] {message}";
        }
    }
}