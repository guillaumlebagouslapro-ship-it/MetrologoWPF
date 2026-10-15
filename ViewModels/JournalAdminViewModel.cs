using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Services.Journal;
using System.Collections.ObjectModel;
using System.Diagnostics;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// Journal d'audit administrateur, présenté par CONNEXION : qui s'est connecté (quand, sur quel
    /// poste) puis les actions de configuration faites pendant cette connexion — ou « connexion
    /// seule » s'il n'a rien modifié. Les refus d'accès sont regroupés, les mises à jour automatiques
    /// (Besançon) masquées par défaut. Source : le fichier partagé <see cref="JournalAdminService.Chemin"/>,
    /// regroupé par <see cref="JournalAdminRegroupement"/>.
    /// </summary>
    public partial class JournalAdminViewModel : ObservableObject
    {
        private List<BlocJournalAdmin> _tous = new();

        [ObservableProperty] private ObservableCollection<BlocJournalAdmin> _blocs = new();
        [ObservableProperty] private string _recherche = string.Empty;
        [ObservableProperty] private bool _afficherAutomatiques;
        [ObservableProperty] private string _statut = "Prêt.";

        public string Chemin => JournalAdminService.Chemin;

        public JournalAdminViewModel()
        {
            _ = RafraichirAsync();
        }

        partial void OnRechercheChanged(string value) => AppliquerFiltre();
        partial void OnAfficherAutomatiquesChanged(bool value) => AppliquerFiltre();

        [RelayCommand]
        private async Task RafraichirAsync()
        {
            Statut = "Chargement…";
            try
            {
                _tous = await Task.Run(() =>
                    JournalAdminRegroupement.Regrouper(JournalAdminService.LireChronologique()));
                AppliquerFiltre();
            }
            catch (Exception ex)
            {
                Statut = $"Lecture échouée : {ex.Message}";
            }
        }

        private void AppliquerFiltre()
        {
            IEnumerable<BlocJournalAdmin> src = _tous;
            if (!AfficherAutomatiques) src = src.Where(b => b.Type != TypeBlocAdmin.Automatique);

            string q = (Recherche ?? string.Empty).Trim();
            if (q.Length > 0) src = src.Where(b => b.Contient(q));

            var liste = src.ToList();
            Blocs = new ObservableCollection<BlocJournalAdmin>(liste);

            int connexions = liste.Count(b => b.Type == TypeBlocAdmin.Connexion);
            int actions = liste.Where(b => b.Type != TypeBlocAdmin.Refus).Sum(b => b.Actions.Sum(a => a.Nombre));
            int refus = liste.Where(b => b.Type == TypeBlocAdmin.Refus).Sum(b => b.NbTentatives);
            Statut = $"{connexions} connexion(s) · {actions} action(s)"
                   + (refus > 0 ? $" · {refus} accès refusé(s)" : string.Empty)
                   + $" · fichier : {Chemin}";
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
