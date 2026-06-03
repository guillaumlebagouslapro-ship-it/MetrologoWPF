using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Services.Journal;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// Journal d'activité Admin — vue par fiche d'intervention (FI). Chaque FI possède son propre
    /// fichier log (<c>Journal_&lt;FI&gt;.txt</c>) dans son dossier ; cette vue liste les FI avec
    /// quelques infos (opérateurs, dernière activité, nb de mesures) et un accès rapide pour ouvrir
    /// le fichier log ou son dossier. Remplace l'ancien journal SQL centralisé.
    /// </summary>
    public partial class JournalViewerViewModel : ObservableObject
    {
        private List<FicheJournalInfo> _toutes = new();

        [ObservableProperty] private ObservableCollection<FicheJournalInfo> _fiches = new();
        [ObservableProperty] private string _recherche = string.Empty;
        [ObservableProperty] private string _statut = "Prêt.";

        /// <summary>Dossier racine (Mesures) scanné — affiché en pied de fenêtre.</summary>
        public string DossierRacine => JournalFIListeService.DossierRacine;

        public JournalViewerViewModel()
        {
            _ = RafraichirAsync();
        }

        partial void OnRechercheChanged(string value) => AppliquerFiltre();

        [RelayCommand]
        private async Task RafraichirAsync()
        {
            Statut = "Chargement…";
            try
            {
                _toutes = await Task.Run(() => JournalFIListeService.Lister());
                AppliquerFiltre();
                Statut = $"{_toutes.Count} fiche(s) d'intervention · dossier : {DossierRacine}";
            }
            catch (Exception ex)
            {
                Statut = $"Lecture échouée : {ex.Message}";
            }
        }

        private void AppliquerFiltre()
        {
            IEnumerable<FicheJournalInfo> src = _toutes;
            string q = (Recherche ?? string.Empty).Trim();
            if (q.Length > 0)
            {
                src = src.Where(f =>
                    f.NumFI.Contains(q, StringComparison.OrdinalIgnoreCase)
                    || f.Operateurs.Any(o => o.Contains(q, StringComparison.OrdinalIgnoreCase)));
            }

            Fiches.Clear();
            foreach (var f in src) Fiches.Add(f);
        }

        /// <summary>Ouvre le fichier <c>Journal_&lt;FI&gt;.txt</c> avec l'application par défaut.</summary>
        [RelayCommand]
        private void OuvrirFichier(FicheJournalInfo? fiche)
        {
            if (fiche == null || !fiche.LogPresent)
            {
                Statut = "Aucun fichier journal pour cette FI.";
                return;
            }
            try
            {
                Process.Start(new ProcessStartInfo(fiche.CheminFichierLog) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Statut = $"Ouverture du fichier échouée : {ex.Message}";
            }
        }

        /// <summary>Ouvre l'explorateur sur le dossier de la FI (fichier log sélectionné si présent).</summary>
        [RelayCommand]
        private void OuvrirDossier(FicheJournalInfo? fiche)
        {
            if (fiche == null || string.IsNullOrEmpty(fiche.CheminDossier))
            {
                Statut = "Dossier introuvable.";
                return;
            }
            try
            {
                if (fiche.LogPresent)
                    Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{fiche.CheminFichierLog}\"") { UseShellExecute = true });
                else
                    Process.Start(new ProcessStartInfo(fiche.CheminDossier) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Statut = $"Ouverture du dossier échouée : {ex.Message}";
            }
        }
    }
}
