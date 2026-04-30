using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Services.Journal;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// Représente l'ensemble des sessions ouvertes durant un jour calendaire donné.
    /// Affiché dans la sidebar gauche du JournalViewer pour permettre à l'admin de
    /// naviguer rapidement par date sans scroller dans une liste plate.
    /// </summary>
    public class GroupeJournalParJour
    {
        public DateTime Jour { get; set; }
        public List<SessionJournal> Sessions { get; set; } = new();

        public int NbSessions => Sessions.Count;
        public bool EstAujourdhui => Jour.Date == DateTime.Today;
        public bool EstHier => Jour.Date == DateTime.Today.AddDays(-1);

        /// <summary>Libellé adaptatif : « Aujourd'hui », « Hier », ou « lundi 28 avril ».</summary>
        public string Libelle
        {
            get
            {
                if (EstAujourdhui) return "Aujourd'hui";
                if (EstHier) return "Hier";
                // « lundi 28 avril » (français long), capitalisation manuelle.
                var lib = Jour.ToString("dddd d MMMM", new CultureInfo("fr-FR"));
                return char.ToUpper(lib[0]) + lib[1..];
            }
        }

        public string DateCourte => Jour.ToString("dd/MM");
        public string DateLongue => Jour.ToString("dd MMMM yyyy", new CultureInfo("fr-FR"));

        /// <summary>Vrai si au moins une session a une erreur — pour pastille rouge dans la sidebar.</summary>
        public bool ContientErreurs => Sessions.Any(s => s.HasErreurs);
        public int TotalErreurs => Sessions.Sum(s => s.NbErreurs);
    }

    /// <summary>
    /// Mois archivé sur disque (lecture transparente depuis JSON via <see cref="ArchivesLogsService"/>).
    /// Affiché en bas de la sidebar, après les jours en base.
    /// </summary>
    public partial class MoisArchive : ObservableObject
    {
        public DateTime Mois { get; set; }

        [ObservableProperty] private bool _estDeplie;
        public ObservableCollection<JourArchive> Jours { get; } = new();

        public string Libelle => Mois.ToString("MMMM yyyy", new CultureInfo("fr-FR"));
        public string LibelleCapitalise =>
            char.ToUpper(Libelle[0]) + Libelle[1..];
    }

    public class JourArchive
    {
        public DateTime Jour { get; set; }
        public string Libelle => Jour.ToString("dddd d MMMM", new CultureInfo("fr-FR"));
        public string LibelleCapitalise =>
            char.ToUpper(Libelle[0]) + Libelle[1..];
        public string DateCourte => Jour.ToString("dd/MM");
    }

    public partial class JournalViewerViewModel : ObservableObject
    {
        // -------------------------------------------------------------------------
        // Données : logs en base + archives sur disque
        // -------------------------------------------------------------------------

        [ObservableProperty] private ObservableCollection<GroupeJournalParJour> _groupesParJour = new();

        /// <summary>Mois archivés sur disque, listés en bas de la sidebar.</summary>
        [ObservableProperty] private ObservableCollection<MoisArchive> _moisArchives = new();

        /// <summary>
        /// Jour archivé actuellement consulté (null = on est sur des logs en base).
        /// Quand non-null, <see cref="SessionsAffichees"/> est rempli depuis le JSON.
        /// </summary>
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(SessionsAffichees))]
        [NotifyPropertyChangedFor(nameof(TitreVue))]
        [NotifyPropertyChangedFor(nameof(EstVueArchivee))]
        private JourArchive? _jourArchiveSelectionne;

        /// <summary>Sessions chargées depuis JSON quand on consulte un jour archivé.</summary>
        private List<SessionJournal> _sessionsArchivees = new();

        public bool EstVueArchivee => JourArchiveSelectionne != null;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(SessionsAffichees))]
        [NotifyPropertyChangedFor(nameof(TitreVue))]
        private GroupeJournalParJour? _jourSelectionne;

        /// <summary>
        /// Sessions à afficher dans le panneau principal. Si un jour est sélectionné
        /// dans la sidebar → ses sessions ; sinon (aucun jour) → toutes les sessions
        /// chargées (vue agrégée).
        /// </summary>
        public ObservableCollection<SessionJournal> SessionsAffichees
        {
            get
            {
                var col = new ObservableCollection<SessionJournal>();

                // Vue archivée : sessions chargées depuis JSON (priorité sur la base)
                if (JourArchiveSelectionne != null)
                {
                    foreach (var s in _sessionsArchivees) col.Add(s);
                    return col;
                }

                if (JourSelectionne != null)
                {
                    foreach (var s in JourSelectionne.Sessions) col.Add(s);
                }
                else
                {
                    foreach (var g in GroupesParJour)
                        foreach (var s in g.Sessions) col.Add(s);
                }
                return col;
            }
        }

        public string TitreVue
        {
            get
            {
                if (JourArchiveSelectionne != null)
                    return $"📦 {JourArchiveSelectionne.LibelleCapitalise} (archivé)";
                return JourSelectionne?.Libelle ?? "Toutes les sessions";
            }
        }

        // -------------------------------------------------------------------------
        // Filtres
        // -------------------------------------------------------------------------

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
        /// technique caché par défaut.
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

        partial void OnAfficherDetailsChanged(bool value)
        {
            // Re-query la base pour appliquer le pré-filtre métier côté SQL.
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

        // -------------------------------------------------------------------------
        // Filtres rapides (boutons Aujourd'hui / 7 j / 30 j / Erreurs / Tout)
        // -------------------------------------------------------------------------

        public enum PeriodeRapide { Aujourdhui, SeptJours, TrenteJours, Tout, ErreursSeulement }

        [ObservableProperty]
        private PeriodeRapide _periodeActive = PeriodeRapide.SeptJours;

        public bool PeriodeAujourdhui => PeriodeActive == PeriodeRapide.Aujourdhui;
        public bool Periode7Jours => PeriodeActive == PeriodeRapide.SeptJours;
        public bool Periode30Jours => PeriodeActive == PeriodeRapide.TrenteJours;
        public bool PeriodeTout => PeriodeActive == PeriodeRapide.Tout;
        public bool PeriodeErreurs => PeriodeActive == PeriodeRapide.ErreursSeulement;

        partial void OnPeriodeActiveChanged(PeriodeRapide value)
        {
            OnPropertyChanged(nameof(PeriodeAujourdhui));
            OnPropertyChanged(nameof(Periode7Jours));
            OnPropertyChanged(nameof(Periode30Jours));
            OnPropertyChanged(nameof(PeriodeTout));
            OnPropertyChanged(nameof(PeriodeErreurs));
        }

        [RelayCommand]
        private async Task FiltrerAujourdhuiAsync()
        {
            PeriodeActive = PeriodeRapide.Aujourdhui;
            FiltreDepuis = DateTime.Today;
            FiltreSeveriteMin = null;
            await RafraichirAsync();
        }

        [RelayCommand]
        private async Task Filtrer7JoursAsync()
        {
            PeriodeActive = PeriodeRapide.SeptJours;
            FiltreDepuis = DateTime.Today.AddDays(-7);
            FiltreSeveriteMin = null;
            await RafraichirAsync();
        }

        [RelayCommand]
        private async Task Filtrer30JoursAsync()
        {
            PeriodeActive = PeriodeRapide.TrenteJours;
            FiltreDepuis = DateTime.Today.AddDays(-30);
            FiltreSeveriteMin = null;
            await RafraichirAsync();
        }

        [RelayCommand]
        private async Task FiltrerToutAsync()
        {
            PeriodeActive = PeriodeRapide.Tout;
            // 5 ans en arrière = pratiquement « tout » sans MIN(DateTime).
            FiltreDepuis = DateTime.Today.AddYears(-5);
            FiltreSeveriteMin = null;
            await RafraichirAsync();
        }

        [RelayCommand]
        private async Task FiltrerErreursSeulementAsync()
        {
            PeriodeActive = PeriodeRapide.ErreursSeulement;
            FiltreDepuis = DateTime.Today.AddDays(-30);
            FiltreSeveriteMin = SeveriteLog.Avertissement;
            await RafraichirAsync();
        }

        // -------------------------------------------------------------------------
        // Cycle de vie
        // -------------------------------------------------------------------------

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
                ActionsMetier = AfficherDetails ? null : _actionsMetier.ToList()
            };

            var sessions = await Journal.Service.ChargerSessionsAsync(filtre);

            // Regroupement par jour calendaire (heure locale) — ordre antichronologique
            // pour avoir « Aujourd'hui » en haut de la sidebar.
            var groupes = sessions
                .GroupBy(s => s.Debut.Date)
                .Select(g => new GroupeJournalParJour
                {
                    Jour = g.Key,
                    Sessions = g.OrderByDescending(s => s.Debut).ToList()
                })
                .OrderByDescending(g => g.Jour)
                .ToList();

            GroupesParJour.Clear();
            foreach (var g in groupes) GroupesParJour.Add(g);

            // Présélection : aujourd'hui si présent, sinon le jour le plus récent,
            // sinon « tout » (null) pour montrer un état clair quand vide.
            JourSelectionne = GroupesParJour.FirstOrDefault(g => g.EstAujourdhui)
                              ?? GroupesParJour.FirstOrDefault();

            // Liste des utilisateurs pour le filtre
            var users = await Journal.Service.ChargerListeUtilisateursAsync();
            Utilisateurs.Clear();
            Utilisateurs.Add(string.Empty);
            foreach (var u in users) Utilisateurs.Add(u);

            int totalSessions = sessions.Count;
            int totalEntries = sessions.Sum(s => s.Entrees.Count);
            Statut = AfficherDetails
                ? $"{totalSessions} session(s) · {totalEntries} entrée(s) (mode debug)."
                : $"{totalSessions} session(s) · {totalEntries} entrée(s) métier.";

            // Charger la liste des mois archivés (lecture FS, pas de SQL)
            ChargerMoisArchives();

            // Notifier les vues que SessionsAffichees doit être recalculé
            OnPropertyChanged(nameof(SessionsAffichees));
            OnPropertyChanged(nameof(TitreVue));
        }

        private void ChargerMoisArchives()
        {
            MoisArchives.Clear();
            foreach (var mois in ArchivesLogsService.ListerMoisArchives())
            {
                var ma = new MoisArchive { Mois = mois };
                foreach (var jour in ArchivesLogsService.ListerJoursDuMoisArchive(mois))
                    ma.Jours.Add(new JourArchive { Jour = jour });
                MoisArchives.Add(ma);
            }
        }

        [RelayCommand]
        private async Task SelectionnerJourArchiveAsync(JourArchive jour)
        {
            if (jour == null) return;
            JourSelectionne = null;  // désélectionne la vue base
            JourArchiveSelectionne = jour;

            try
            {
                Statut = $"Chargement archive du {jour.Jour:dd/MM/yyyy}…";
                _sessionsArchivees = await ArchivesLogsService.ChargerJournauxArchivesAsync(jour.Jour);
                Statut = $"Archive {jour.Jour:dd/MM/yyyy} : {_sessionsArchivees.Count} session(s).";
                OnPropertyChanged(nameof(SessionsAffichees));
            }
            catch (Exception ex)
            {
                Statut = $"Lecture archive échouée : {ex.Message}";
            }
        }

        [RelayCommand]
        private void BasculerExpandMoisArchive(MoisArchive mois)
        {
            if (mois != null) mois.EstDeplie = !mois.EstDeplie;
        }

        [RelayCommand] private async Task AppliquerFiltreAsync() => await RafraichirAsync();

        /// <summary>
        /// Force l'archivage du mois précédent (export JSON + suppression SQL). À utiliser
        /// pour tester sans attendre le 1er du mois, ou pour rattraper un mois manqué.
        /// </summary>
        [RelayCommand]
        private async Task ArchiverMaintenantAsync()
        {
            var maintenant = DateTime.Now;
            var moisPrec = new DateTime(maintenant.Year, maintenant.Month, 1).AddMonths(-1);

            var conf = System.Windows.MessageBox.Show(
                $"Archiver les logs de {moisPrec:MMMM yyyy} maintenant ?\n\n" +
                "Les logs seront exportés en JSON puis supprimés de la base SQL Server.\n" +
                $"Dossier : {ArchivesLogsService.DossierArchivesRacine}",
                "Archiver les logs",
                System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question);
            if (conf != System.Windows.MessageBoxResult.Yes) return;

            try
            {
                Statut = "Archivage en cours…";
                int n = await ArchivesLogsService.ArchiverMoisAsync(moisPrec, force: true);
                Statut = $"Archivage OK : {n} entrée(s) exportée(s) dans {moisPrec:yyyy-MM}.";
                await RafraichirAsync();
            }
            catch (Exception ex)
            {
                Statut = $"Archivage échoué : {ex.Message}";
            }
        }

        [RelayCommand]
        private void SelectionnerTousLesJours()
        {
            JourSelectionne = null;
            JourArchiveSelectionne = null;
            OnPropertyChanged(nameof(SessionsAffichees));
            OnPropertyChanged(nameof(TitreVue));
        }

        partial void OnJourSelectionneChanged(GroupeJournalParJour? value)
        {
            // Sélectionner un jour en base désélectionne le jour archivé éventuel.
            if (value != null && JourArchiveSelectionne != null)
                JourArchiveSelectionne = null;
        }
    }
}
