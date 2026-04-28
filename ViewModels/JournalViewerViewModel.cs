using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Services.Journal;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
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

        /// <summary>
        /// Mode debug : si décoché (défaut), seules les entrées « métier » sont affichées
        /// (sessions, configurations validées, mesures démarrées/terminées/échouées,
        /// changements de rubidium, erreurs et avertissements). Si coché, on voit
        /// l'intégralité des entrées techniques (SCPI, GPIB, Excel, perf…).
        /// </summary>
        [ObservableProperty] private bool _afficherDetails;

        /// <summary>
        /// Liste blanche des actions métier — ce qui répond à « qui s'est connecté quand,
        /// avec quel appareil, sur quel FI, quel type de mesure ». Le reste est du bruit
        /// technique.
        /// </summary>
        private static readonly HashSet<string> _actionsMetier = new(StringComparer.OrdinalIgnoreCase)
        {
            "DEBUT_SESSION",
            "FIN_SESSION",
            "CONFIG_VALIDEE",
            "MESURE_DEBUT",
            "MESURE_FIN",
            "MESURE_ECHEC",
            "MESURE_STOP",
            "SELECTION_RUBIDIUM",
        };

        /// <summary>
        /// Décide si une entrée doit apparaître dans la vue métier (mode normal).
        /// En mode debug, tout passe.
        /// </summary>
        private bool EstAffichable(LogEntry entry)
        {
            if (AfficherDetails) return true;
            // Erreurs et avertissements toujours visibles : c'est ce qui aide à diagnostiquer
            // un problème sans avoir à activer le mode debug.
            if (entry.Severite == SeveriteLog.Erreur ||
                entry.Severite == SeveriteLog.Avertissement) return true;
            return _actionsMetier.Contains(entry.Action);
        }

        partial void OnAfficherDetailsChanged(bool value)
        {
            // Re-filtre la vue sans re-query la base de données.
            _ = RafraichirAsync();
        }

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
                Recherche = string.IsNullOrEmpty(FiltreRecherche) ? null : FiltreRecherche,
                // Pré-filtre côté SQL : en mode normal, restreint à la liste blanche métier
                // (+ erreurs/avertissements toujours inclus côté service). En mode debug,
                // null = aucune restriction.
                ActionsMetier = AfficherDetails ? null : _actionsMetier.ToList()
            };

            var sessions = await Journal.Service.ChargerSessionsAsync(filtre);

            int totalAffiche = 0;
            foreach (var s in sessions) totalAffiche += s.Entrees.Count;

            Sessions.Clear();
            foreach (var s in sessions) Sessions.Add(s);

            var users = await Journal.Service.ChargerListeUtilisateursAsync();
            Utilisateurs.Clear();
            Utilisateurs.Add(string.Empty); // "Tous"
            foreach (var u in users) Utilisateurs.Add(u);

            Statut = AfficherDetails
                ? $"{sessions.Count} session(s) · {totalAffiche} entrée(s) (mode debug — tout affiché)."
                : $"{sessions.Count} session(s) · {totalAffiche} entrée(s) métier.";
        }

        [RelayCommand] private async Task AppliquerFiltreAsync() => await RafraichirAsync();
    }
}
