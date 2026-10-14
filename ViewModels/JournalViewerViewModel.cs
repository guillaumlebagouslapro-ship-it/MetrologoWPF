using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Services;
using Metrologo.Services.Journal;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace Metrologo.ViewModels
{
    /// <summary>Un mois de logs système, actif (dans Logs/) ou archivé (dans Logs/Archives/YYYY-MM/).</summary>
    public sealed class MoisLogsInfo
    {
        public DateTime Mois { get; set; }
        public bool EstArchive { get; set; }
        public bool EstMoisCourant { get; set; }
        public string Dossier { get; set; } = string.Empty;
        public string CheminTexte { get; set; } = string.Empty;
        public int NbFichiers { get; set; }
        public long TailleOctets { get; set; }

        public string Libelle => Mois.ToString("MMMM yyyy");
        public string EtatAffiche => EstArchive ? "Archivé" : EstMoisCourant ? "En cours" : "À archiver";
        public string TailleAffichee => TailleOctets switch
        {
            < 1024 => $"{TailleOctets} o",
            < 1024 * 1024 => $"{TailleOctets / 1024.0:0.#} Ko",
            _ => $"{TailleOctets / (1024.0 * 1024):0.#} Mo",
        };
        /// <summary>Un mois passé encore présent dans Logs/ (archivage raté) : on peut le rattraper à la main.</summary>
        public bool PeutArchiver => !EstArchive && !EstMoisCourant;
    }

    /// <summary>
    /// Journal &amp; maintenance (zone Admin), en trois onglets :
    /// <list type="bullet">
    ///   <item>Fiches d'intervention : une ligne par FI (partage Mesures + FI restées en local hors
    ///         réseau), avec accès direct à son <c>Journal_&lt;FI&gt;.txt</c> ;</item>
    ///   <item>Journal système : les entrées <c>entries-YYYY-MM.jsonl</c> du dossier Logs (connexions,
    ///         mesures, erreurs, Besançon, etc.), filtrables par période, gravité et texte ;</item>
    ///   <item>Maintenance : les mois de logs (actifs / archivés), l'archivage manuel et l'accès
    ///         au dossier Logs.</item>
    /// </list>
    /// Les onglets Journal système et Maintenance ne sont chargés qu'à leur première ouverture.
    /// </summary>
    public partial class JournalViewerViewModel : ObservableObject
    {
        // =====================================================================
        // Onglet 1 : fiches d'intervention
        // =====================================================================

        private List<FicheJournalInfo> _toutes = new();

        [ObservableProperty] private ObservableCollection<FicheJournalInfo> _fiches = new();
        [ObservableProperty] private string _recherche = string.Empty;
        [ObservableProperty] private string _statut = "Prêt.";
        [ObservableProperty] private int _ongletIndex;

        /// <summary>Dossier racine (Mesures) qu'on parcourt ; rappelé en pied de fenêtre.</summary>
        public string DossierRacine => JournalFIListeService.DossierRacine;

        public JournalViewerViewModel()
        {
            _ = RafraichirAsync();
        }

        partial void OnRechercheChanged(string value) => AppliquerFiltre();

        partial void OnOngletIndexChanged(int value)
        {
            if (value == 1 && !_systemeCharge) _ = RafraichirSystemeAsync();
            if (value == 2 && !_maintenanceChargee) RafraichirMaintenance();
        }

        [RelayCommand]
        private async Task RafraichirAsync()
        {
            Statut = "Chargement…";
            try
            {
                bool reseauOk = await Task.Run(JournalFIListeService.RacineAccessible);
                _toutes = await Task.Run(() => JournalFIListeService.Lister());
                AppliquerFiltre();

                int locales = _toutes.Count(f => f.EstLocale);
                string detailLocales = locales > 0 ? $" (dont {locales} en local, pas encore sur le réseau)" : string.Empty;
                Statut = reseauOk
                    ? $"{_toutes.Count} fiche(s) d'intervention{detailLocales} · dossier : {DossierRacine}"
                    : $"⚠ Dossier Mesures inaccessible ({DossierRacine}) · {_toutes.Count} fiche(s) locale(s) seulement";
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

        /// <summary>Ouvre le fichier <c>Journal_&lt;FI&gt;.txt</c> dans l'application associée par Windows.</summary>
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

        /// <summary>Ouvre l'explorateur sur le dossier de la FI, en pré-sélectionnant le fichier log s'il existe.</summary>
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

        // =====================================================================
        // Onglet 2 : journal système (entries-YYYY-MM.jsonl)
        // =====================================================================

        /// <summary>Au-delà, on n'affiche que les plus récentes (la liste reste fluide).</summary>
        private const int MaxEntreesAffichees = 5000;

        private List<LogEntry> _toutesEntrees = new();
        private bool _systemeCharge;

        public IReadOnlyList<string> Periodes { get; } =
            new[] { "Aujourd'hui", "7 derniers jours", "30 derniers jours", "Tout" };

        [ObservableProperty] private ObservableCollection<LogEntry> _entreesSysteme = new();
        [ObservableProperty] private string _periodeSysteme = "7 derniers jours";
        [ObservableProperty] private bool _anomaliesSeulement;
        [ObservableProperty] private string _rechercheSysteme = string.Empty;
        [ObservableProperty] private string _statutSysteme = string.Empty;
        [ObservableProperty] private LogEntry? _entreeSelectionnee;

        partial void OnPeriodeSystemeChanged(string value) => _ = RafraichirSystemeAsync();
        partial void OnAnomaliesSeulementChanged(bool value) => _ = RafraichirSystemeAsync();
        partial void OnRechercheSystemeChanged(string value) => AppliquerFiltreSysteme();

        private DateTime? DebutPeriode() => PeriodeSysteme switch
        {
            "Aujourd'hui" => DateTime.Today,
            "7 derniers jours" => DateTime.Today.AddDays(-7),
            "30 derniers jours" => DateTime.Today.AddDays(-30),
            _ => null,
        };

        [RelayCommand]
        private async Task RafraichirSystemeAsync()
        {
            _systemeCharge = true;
            var service = Journal.Service;
            if (service == null)
            {
                StatutSysteme = "Journal non configuré.";
                return;
            }

            StatutSysteme = "Chargement…";
            try
            {
                var filtre = new FiltreJournal
                {
                    Depuis = DebutPeriode(),
                    SeveriteMin = AnomaliesSeulement ? SeveriteLog.Avertissement : null,
                };
                _toutesEntrees = await service.ChargerEntreesAsync(filtre);
                AppliquerFiltreSysteme();
            }
            catch (Exception ex)
            {
                StatutSysteme = $"Lecture échouée : {ex.Message}";
            }
        }

        private void AppliquerFiltreSysteme()
        {
            IEnumerable<LogEntry> src = _toutesEntrees;
            string q = (RechercheSysteme ?? string.Empty).Trim();
            if (q.Length > 0)
            {
                src = src.Where(e =>
                    (e.Message?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
                    || (e.Action?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
                    || (e.Utilisateur?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
                    || (e.Machine?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
                    || e.CategorieLibelle.Contains(q, StringComparison.OrdinalIgnoreCase));
            }

            var filtrees = src.ToList();
            EntreesSysteme = new ObservableCollection<LogEntry>(filtrees.Take(MaxEntreesAffichees));

            int nbErr = filtrees.Count(e => e.Severite == SeveriteLog.Erreur);
            int nbWarn = filtrees.Count(e => e.Severite == SeveriteLog.Avertissement);
            string tronque = filtrees.Count > MaxEntreesAffichees
                ? $" (les {MaxEntreesAffichees} plus récentes affichées)" : string.Empty;

            if (!Directory.Exists(CheminsMetrologo.Logs))
                StatutSysteme = $"⚠ Dossier Logs inaccessible : {CheminsMetrologo.Logs}";
            else
                StatutSysteme = $"{filtrees.Count} entrée(s){tronque} · {nbErr} erreur(s) · {nbWarn} avertissement(s) · dossier : {CheminsMetrologo.Logs}";
        }

        /// <summary>Ouvre le TXT lisible du mois de l'entrée sélectionnée (ou du mois courant).</summary>
        [RelayCommand]
        private void OuvrirTexteDuMois()
        {
            var mois = EntreeSelectionnee?.Timestamp ?? DateTime.Now;
            string nom = $"entries-{mois:yyyy-MM}.txt";
            string actif = Path.Combine(CheminsMetrologo.Logs, nom);
            string archive = Path.Combine(ArchivesLogsService.DossierArchivesRacine, mois.ToString("yyyy-MM"), nom);
            string? chemin = File.Exists(actif) ? actif : File.Exists(archive) ? archive : null;
            if (chemin == null)
            {
                StatutSysteme = $"Aucun fichier texte pour {mois:MMMM yyyy}.";
                return;
            }
            OuvrirAvecWindows(chemin, s => StatutSysteme = s);
        }

        // =====================================================================
        // Onglet 3 : maintenance
        // =====================================================================

        private bool _maintenanceChargee;

        [ObservableProperty] private ObservableCollection<MoisLogsInfo> _moisLogs = new();
        [ObservableProperty] private string _statutMaintenance = string.Empty;
        [ObservableProperty] private string _etatDossierLogs = string.Empty;
        [ObservableProperty] private string _etatDossierMesures = string.Empty;
        [ObservableProperty] private string _etatAuditAdmin = string.Empty;

        public string DossierLogs => CheminsMetrologo.Logs;
        public string DossierArchives => ArchivesLogsService.DossierArchivesRacine;

        [RelayCommand]
        private void RafraichirMaintenance()
        {
            _maintenanceChargee = true;
            try
            {
                EtatDossierLogs = Directory.Exists(DossierLogs)
                    ? $"✔ Accessible — {DossierLogs}"
                    : $"✖ Inaccessible — {DossierLogs}";
                EtatDossierMesures = JournalFIListeService.RacineAccessible()
                    ? $"✔ Accessible — {DossierRacine}"
                    : $"✖ Inaccessible — {DossierRacine}";
                EtatAuditAdmin = File.Exists(JournalAdminService.Chemin)
                    ? $"✔ {JournalAdminService.Chemin} ({new FileInfo(JournalAdminService.Chemin).Length / 1024.0:0.#} Ko)"
                    : $"— Pas encore créé : {JournalAdminService.Chemin}";

                var liste = ListerMoisLogs();
                MoisLogs = new ObservableCollection<MoisLogsInfo>(liste);

                int aArchiver = liste.Count(m => m.PeutArchiver);
                StatutMaintenance = aArchiver > 0
                    ? $"⚠ {aArchiver} mois passé(s) encore dans Logs : archivage automatique non fait (bouton « Archiver »)."
                    : $"{liste.Count} mois de logs · archivage automatique au premier lancement de chaque mois.";
            }
            catch (Exception ex)
            {
                StatutMaintenance = $"Lecture échouée : {ex.Message}";
            }
        }

        private static List<MoisLogsInfo> ListerMoisLogs()
        {
            var parMois = new Dictionary<DateTime, MoisLogsInfo>();
            var moisCourant = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            string logs = CheminsMetrologo.Logs;

            void Ajouter(string dossier, bool archive)
            {
                foreach (var f in Directory.EnumerateFiles(dossier, "entries-*.*"))
                {
                    string nom = Path.GetFileNameWithoutExtension(f);
                    if (!DateTime.TryParseExact(nom.Substring("entries-".Length), "yyyy-MM",
                            System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.None, out var mois))
                        continue;

                    if (!parMois.TryGetValue(mois, out var info) || (info.EstArchive && !archive))
                    {
                        // Si un mois traîne à la fois dans Logs/ et dans Archives/, c'est la copie
                        // active qui compte : elle signale un archivage incomplet.
                        info = new MoisLogsInfo
                        {
                            Mois = mois,
                            EstArchive = archive,
                            EstMoisCourant = mois == moisCourant,
                            Dossier = dossier,
                        };
                        parMois[mois] = info;
                    }
                    if (info.EstArchive != archive) continue;

                    info.NbFichiers++;
                    try { info.TailleOctets += new FileInfo(f).Length; } catch { }
                    if (f.EndsWith(".txt", StringComparison.OrdinalIgnoreCase)) info.CheminTexte = f;
                }
            }

            if (Directory.Exists(logs)) Ajouter(logs, archive: false);
            string archives = ArchivesLogsService.DossierArchivesRacine;
            if (Directory.Exists(archives))
                foreach (var sd in Directory.EnumerateDirectories(archives))
                    Ajouter(sd, archive: true);

            return parMois.Values.OrderByDescending(m => m.Mois).ToList();
        }

        [RelayCommand]
        private void OuvrirDossierLogs()
        {
            if (!Directory.Exists(DossierLogs))
            {
                StatutMaintenance = $"Dossier Logs inaccessible : {DossierLogs}";
                return;
            }
            OuvrirAvecWindows(DossierLogs, s => StatutMaintenance = s);
        }

        [RelayCommand]
        private void OuvrirMoisLogs(MoisLogsInfo? mois)
        {
            if (mois == null) return;
            string cible = !string.IsNullOrEmpty(mois.CheminTexte) && File.Exists(mois.CheminTexte)
                ? mois.CheminTexte
                : mois.Dossier;
            OuvrirAvecWindows(cible, s => StatutMaintenance = s);
        }

        /// <summary>
        /// Range tout de suite un mois passé dans Logs/Archives/YYYY-MM/, sans attendre le prochain
        /// démarrage au 1er du mois. Sert à rattraper un mois resté en plan (PC éteint, partage
        /// injoignable au moment du démarrage…).
        /// </summary>
        [RelayCommand]
        private async Task ArchiverMoisAsync(MoisLogsInfo? mois)
        {
            if (mois == null || !mois.PeutArchiver) return;

            var conf = MessageBox.Show(
                $"Archiver les logs de {mois.Libelle} maintenant ?\n\n"
              + $"Les fichiers entries-{mois.Mois:yyyy-MM}.* seront déplacés vers :\n"
              + $"{Path.Combine(DossierArchives, mois.Mois.ToString("yyyy-MM"))}\n\n"
              + "Ils restent consultables depuis l'onglet Journal système.",
                "Archiver les logs", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (conf != MessageBoxResult.Yes) return;

            try
            {
                int n = await Task.Run(() => ArchivesLogsService.ArchiverMoisAsync(mois.Mois, force: true));
                RafraichirMaintenance();
                StatutMaintenance = n > 0
                    ? $"Archivage de {mois.Libelle} terminé : {n} fichier(s) déplacé(s)."
                    : $"Rien n'a été déplacé pour {mois.Libelle} (déjà archivé, ou archivage en cours sur un autre poste).";
                _systemeCharge = false; // les chemins ont changé : on relira à la prochaine ouverture
            }
            catch (Exception ex)
            {
                StatutMaintenance = $"Archivage échoué : {ex.Message}";
            }
        }

        // =====================================================================

        private static void OuvrirAvecWindows(string chemin, Action<string> signalerErreur)
        {
            try
            {
                Process.Start(new ProcessStartInfo(chemin) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                signalerErreur($"Ouverture échouée : {ex.Message}");
            }
        }
    }
}
