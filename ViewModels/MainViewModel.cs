using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services.Journal;
using Metrologo.Views;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace Metrologo.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(TexteUtilisateurConnecte))]
        private Utilisateur? _utilisateurConnecte;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(EstEnModeAdmin))]
        [NotifyPropertyChangedFor(nameof(EstEnSelectionUtilisateur))]
        [NotifyPropertyChangedFor(nameof(EstEnSelectionPoste))]
        [NotifyPropertyChangedFor(nameof(NavigationActive))]
        [NotifyPropertyChangedFor(nameof(TexteMode))]
        [NotifyPropertyChangedFor(nameof(TexteUtilisateurConnecte))]
        private object? _vueActuelle;

        [ObservableProperty]
        private string _titreApplication = "Metrologo v2026";

        [ObservableProperty]
        private bool _estSurBaie;

        private readonly AccueilViewModel _accueilViewModel = new();

        /// <summary>
        /// Expose le ViewModel d'accueil pour que le bandeau de navigation
        /// (MainWindow.xaml) puisse y binder les commandes/propriétés GPIB qui ont été
        /// remontées hors de la zone de mesure (scan, badge nb appareils, etc.).
        /// </summary>
        public AccueilViewModel Accueil => _accueilViewModel;
        private readonly AdminViewModel _adminViewModel = new();
        private readonly SelectionPosteViewModel _selectionPosteViewModel = new();
        private readonly SelectionUtilisateurViewModel _selectionUtilisateurViewModel = new();

        public bool EstEnModeAdmin => VueActuelle is AdminViewModel;
        public bool EstEnSelectionPoste => VueActuelle is SelectionPosteViewModel;
        public bool EstEnSelectionUtilisateur => VueActuelle is SelectionUtilisateurViewModel;

        /// <summary>
        /// Faux tant que l'utilisateur n'a pas choisi son identité + son poste : la barre
        /// de navigation et la barre de statut sont alors masquées pour éviter qu'on
        /// contourne ces étapes en cliquant sur « Accueil ».
        /// </summary>
        public bool NavigationActive => !EstEnSelectionPoste && !EstEnSelectionUtilisateur;

        public string TexteMode => EstEnModeAdmin ? "Mode : Administration" : "Mode : Exploitation";

        public string TexteUtilisateurConnecte
        {
            get
            {
                if (UtilisateurConnecte == null) return "Utilisateur : non connecté";
                if (EstEnSelectionUtilisateur) return string.Empty;
                if (EstEnSelectionPoste) return $"Utilisateur : {UtilisateurConnecte.NomComplet}";
                return $"Utilisateur : {UtilisateurConnecte.NomComplet} ({(EstSurBaie ? "Baie" : "Paillasse")})";
            }
        }
        public string RubidiumActifTexte => _accueilViewModel.RubidiumActifTexte;

        // ---- Changements admin reçus depuis un AUTRE poste (indicateur persistant ⚠) ----
        private readonly List<EntreeJournalAdmin> _changementsAdmin = new();

        /// <summary>Vrai dès qu'au moins un changement admin a été reçu et pas encore acquitté
        /// (affiche le triangle d'alerte dans le bandeau ; disparaît au clic).</summary>
        [ObservableProperty] private bool _aChangementsAdminEnAttente;

        /// <summary>Nombre de changements non acquittés (badge sur le triangle).</summary>
        [ObservableProperty] private int _nbChangementsAdmin;

        /// <summary>Résumé pour l'infobulle du triangle.</summary>
        public string ResumeChangementsAdmin =>
            _changementsAdmin.Count == 0
                ? string.Empty
                : "Changements de configuration récents :\n"
                  + string.Join("\n", _changementsAdmin
                        .Skip(System.Math.Max(0, _changementsAdmin.Count - 8))
                        .Select(e => $"• {e.Horodatage:HH:mm} — {e.ActionLisible}"
                                   + (string.IsNullOrWhiteSpace(e.Utilisateur) ? "" : $" (par {e.Utilisateur})")));

        private void OnChangementsAdminRecus(IReadOnlyList<EntreeJournalAdmin> nouveaux)
        {
            _changementsAdmin.AddRange(nouveaux);
            if (_changementsAdmin.Count > 100)
                _changementsAdmin.RemoveRange(0, _changementsAdmin.Count - 100);

            NbChangementsAdmin = _changementsAdmin.Count;
            AChangementsAdminEnAttente = true;
            OnPropertyChanged(nameof(ResumeChangementsAdmin));
        }

        /// <summary>Clic sur le triangle : affiche le détail des changements puis efface l'indicateur.</summary>
        [RelayCommand]
        private void AcquitterChangementsAdmin()
        {
            if (_changementsAdmin.Count > 0)
            {
                string liste = string.Join("\n", _changementsAdmin.Select(e =>
                    $"• {e.Horodatage:dd/MM HH:mm} — {e.ActionLisible}"
                  + (string.IsNullOrWhiteSpace(e.Detail) ? "" : $" : {e.Detail}")
                  + (string.IsNullOrWhiteSpace(e.Utilisateur) ? "" : $"  (par {e.Utilisateur})")));
                MessageBox.Show(
                    "Des changements de configuration ont été appliqués depuis un autre poste :\n\n" + liste,
                    "Changements administrateur", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            _changementsAdmin.Clear();
            NbChangementsAdmin = 0;
            AChangementsAdminEnAttente = false;
            OnPropertyChanged(nameof(ResumeChangementsAdmin));
        }

        public MainViewModel()
        {
            NotificationsAdminWatcher.ChangementsRecus += OnChangementsAdminRecus;

            // Étape 1 : sélection de l'utilisateur dans le menu déroulant.
            VueActuelle = _selectionUtilisateurViewModel;

            _selectionUtilisateurViewModel.OnUtilisateurChoisi = (utilisateur) =>
            {
                UtilisateurConnecte = utilisateur;
                EtatApplication.UtilisateurConnecte = utilisateur;
                _ = Metrologo.Services.Journal.Journal.DemarrerSessionAsync(utilisateur.Login);
                // Étape 2 : sélection Baie / Paillasse.
                VueActuelle = _selectionPosteViewModel;
            };

            _selectionPosteViewModel.OnPosteSelectionne = (choixBaie) =>
            {
                EstSurBaie = choixBaie;
                _accueilViewModel.EstSurBaie = choixBaie;
                OnPropertyChanged(nameof(TexteUtilisateurConnecte));
                // Étape 3 : accueil.
                VueActuelle = _accueilViewModel;
                _ = Metrologo.Services.Journal.Journal.DefinirPosteAsync(choixBaie ? "Baie" : "Paillasse");
            };

            // Retour depuis la sélection du poste vers l'écran de connexion (mauvais
            // utilisateur choisi au démarrage) — sans redémarrer l'application.
            _selectionPosteViewModel.OnRetour = () => _ = RevenirAEcranConnexionAsync("Retour sélection poste");

            _accueilViewModel.PropertyChanged += (_, e) =>
            {
                if (e.PropertyName == nameof(AccueilViewModel.RubidiumActifTexte))
                    OnPropertyChanged(nameof(RubidiumActifTexte));
            };
        }

        partial void OnUtilisateurConnecteChanged(Utilisateur? value)
        {
            OnPropertyChanged(nameof(TexteUtilisateurConnecte));
        }

        partial void OnVueActuelleChanged(object? value)
        {
            OnPropertyChanged(nameof(EstEnModeAdmin));
            OnPropertyChanged(nameof(EstEnSelectionPoste));
            OnPropertyChanged(nameof(EstEnSelectionUtilisateur));
            OnPropertyChanged(nameof(TexteMode));
        }

        [RelayCommand]
        private void AllerAccueil() => VueActuelle = _accueilViewModel;

        /// <summary>
        /// Le bouton Administration est visible pour tout le monde. Au clic :
        ///   • Si un admin est DÉJÀ authentifié dans la session courante
        ///     (<see cref="EtatApplication.AdminConnecte"/> non null), on bascule
        ///     directement sur AdminView sans redemander le mot de passe.
        ///   • Sinon, modale d'identifiant + mot de passe. Si OK, on enregistre
        ///     le compte authentifié et on bascule sur AdminView.
        ///
        /// L'authentification reste valide pour toute la durée de l'app (jusqu'à
        /// fermeture du logiciel). Le retour Accueil ne déconnecte plus l'admin
        /// pour éviter de devoir resaisir les identifiants à chaque aller-retour.
        /// </summary>
        [RelayCommand]
        private void OuvrirAdmin()
        {
            // Déjà authentifié dans cette session : raccourci direct.
            if (EtatApplication.AdminConnecte != null)
            {
                VueActuelle = _adminViewModel;
                return;
            }

            var win = new SaisieMdpAdminWindow { Owner = Application.Current.MainWindow };
            if (win.ShowDialog() == true && win.AdminAuthentifie != null)
            {
                EtatApplication.AdminConnecte = win.AdminAuthentifie;
                VueActuelle = _adminViewModel;
            }
        }

        /// <summary>
        /// Retour à l'écran d'accueil. NB : on NE déconnecte PAS l'admin
        /// (<see cref="EtatApplication.AdminConnecte"/> reste actif). L'utilisateur
        /// peut revenir dans la zone admin sans resaisir ses identifiants — la
        /// déconnexion ne se fait qu'à la fermeture de l'app.
        /// </summary>
        [RelayCommand]
        private void RetourAccueil()
        {
            VueActuelle = _accueilViewModel;
        }

        /// <summary>
        /// Revient à l'écran de sélection d'utilisateur sans quitter l'application — pour
        /// changer d'opérateur (relais / reprise de FI), ou prendre en compte un compte tout
        /// juste créé en administration. Clôture proprement la session journal de l'utilisateur
        /// courant, déconnecte l'admin éventuellement authentifié, puis recharge la liste des
        /// comptes (à jour, y compris ceux créés depuis un autre poste) avant de réafficher
        /// l'écran de connexion.
        /// </summary>
        [RelayCommand]
        private Task ChangerUtilisateur() => RevenirAEcranConnexionAsync("Changement d'utilisateur");

        /// <summary>
        /// Logique commune de retour à l'écran de connexion (utilisée par le bouton
        /// « Changer d'utilisateur » de la barre de navigation et par le retour depuis la
        /// sélection du poste). Clôture proprement la session journal de l'utilisateur courant,
        /// déconnecte l'admin éventuellement authentifié, recharge la liste des comptes (à jour,
        /// y compris ceux créés depuis un autre poste) puis réaffiche l'écran de connexion.
        /// </summary>
        private async Task RevenirAEcranConnexionAsync(string motif)
        {
            // Clôture la session journal FI utilisateur + la session système de l'utilisateur
            // courant avant la bascule, pour que la session du prochain utilisateur soit distincte.
            try { Metrologo.Services.Journal.JournalFIService.TerminerSession(motif); }
            catch { /* best-effort */ }
            try { await Metrologo.Services.Journal.Journal.TerminerSessionAsync(); }
            catch { /* best-effort */ }

            // L'admin doit re-saisir ses identifiants après un changement d'utilisateur.
            EtatApplication.AdminConnecte = null;
            EtatApplication.UtilisateurConnecte = null;
            UtilisateurConnecte = null;

            // Recharge la liste à jour (relit le JSON) puis réaffiche l'écran de connexion.
            _selectionUtilisateurViewModel.Recharger();
            VueActuelle = _selectionUtilisateurViewModel;
        }

        [RelayCommand]
        private void Quitter() => Application.Current.Shutdown();
    }
}
