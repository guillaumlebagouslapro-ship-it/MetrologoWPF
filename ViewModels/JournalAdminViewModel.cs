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
    /// Journal d'audit administrateur : liste des actions de configuration (changement de
    /// rubidium, modules d'incertitude, catalogue appareils, utilisateurs, mots de passe…).
    /// Lit le fichier partagé <see cref="JournalAdminService.Chemin"/>. Les consultations
    /// d'écrans / de FI ne sont pas tracées ici.
    /// </summary>
    public partial class JournalAdminViewModel : ObservableObject
    {
        private List<EntreeJournalAdmin> _toutes = new();

        [ObservableProperty] private ObservableCollection<EntreeJournalAdmin> _entrees = new();
        [ObservableProperty] private string _recherche = string.Empty;
        [ObservableProperty] private string _statut = "Prêt.";

        public string Chemin => JournalAdminService.Chemin;

        public JournalAdminViewModel()
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
                _toutes = await Task.Run(() => JournalAdminService.Lire());
                AppliquerFiltre();
                Statut = $"{_toutes.Count} action(s) administrateur · fichier : {Chemin}";
            }
            catch (Exception ex)
            {
                Statut = $"Lecture échouée : {ex.Message}";
            }
        }

        private void AppliquerFiltre()
        {
            IEnumerable<EntreeJournalAdmin> src = _toutes;
            string q = (Recherche ?? string.Empty).Trim();
            if (q.Length > 0)
            {
                src = src.Where(e =>
                    e.Utilisateur.Contains(q, StringComparison.OrdinalIgnoreCase)
                    || e.ActionLisible.Contains(q, StringComparison.OrdinalIgnoreCase)
                    || e.Action.Contains(q, StringComparison.OrdinalIgnoreCase)
                    || e.Detail.Contains(q, StringComparison.OrdinalIgnoreCase));
            }

            Entrees.Clear();
            foreach (var e in src) Entrees.Add(e);
        }

        [RelayCommand]
        private void OuvrirFichier()
        {
            try
            {
                if (System.IO.File.Exists(Chemin))
                    Process.Start(new ProcessStartInfo(Chemin) { UseShellExecute = true });
                else
                    Statut = "Le fichier d'audit n'existe pas encore (aucune action enregistrée).";
            }
            catch (Exception ex)
            {
                Statut = $"Ouverture du fichier échouée : {ex.Message}";
            }
        }
    }
}
