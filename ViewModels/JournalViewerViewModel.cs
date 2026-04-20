using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Services.Journal;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace Metrologo.ViewModels
{
    public partial class JournalViewerViewModel : ObservableObject
    {
        [ObservableProperty] private ObservableCollection<SessionJournal> _sessions = new();
        [ObservableProperty] private ObservableCollection<string> _utilisateurs = new();
        [ObservableProperty] private string? _filtreUtilisateur;
        [ObservableProperty] private DateTime _filtreDepuis = DateTime.Today.AddDays(-7);
        [ObservableProperty] private string _filtreRecherche = string.Empty;
        [ObservableProperty] private CategorieLog? _filtreCategorie;
        [ObservableProperty] private SeveriteLog? _filtreSeveriteMin;
        [ObservableProperty] private string _statut = "Prêt.";

        public IEnumerable<CategorieLog?> CategoriesFiltre
        {
            get
            {
                var list = new List<CategorieLog?> { null };
                foreach (CategorieLog c in Enum.GetValues(typeof(CategorieLog))) list.Add(c);
                return list;
            }
        }

        public IEnumerable<SeveriteLog?> SeveritesFiltre
        {
            get
            {
                var list = new List<SeveriteLog?> { null };
                foreach (SeveriteLog s in Enum.GetValues(typeof(SeveriteLog))) list.Add(s);
                return list;
            }
        }

        public JournalViewerViewModel()
        {
            _ = RafraichirAsync();
        }

        [RelayCommand]
        private async Task RafraichirAsync()
        {
            Statut = "Chargement…";

            if (Journal.Service == null)
            {
                Statut = "Service de journalisation non initialisé.";
                return;
            }

            var filtre = new FiltreJournal
            {
                Depuis = FiltreDepuis,
                Utilisateur = string.IsNullOrEmpty(FiltreUtilisateur) ? null : FiltreUtilisateur,
                Categorie = FiltreCategorie,
                SeveriteMin = FiltreSeveriteMin,
                Recherche = string.IsNullOrEmpty(FiltreRecherche) ? null : FiltreRecherche
            };

            var sessions = await Journal.Service.ChargerSessionsAsync(filtre);
            Sessions.Clear();
            foreach (var s in sessions) Sessions.Add(s);

            var users = await Journal.Service.ChargerListeUtilisateursAsync();
            Utilisateurs.Clear();
            Utilisateurs.Add(string.Empty); // "Tous"
            foreach (var u in users) Utilisateurs.Add(u);

            int totalEntrees = 0;
            foreach (var s in sessions) totalEntrees += s.NbEntrees;
            Statut = $"{sessions.Count} session(s) · {totalEntrees} entrée(s) affichée(s).";
        }

        [RelayCommand] private async Task AppliquerFiltreAsync() => await RafraichirAsync();
    }
}
