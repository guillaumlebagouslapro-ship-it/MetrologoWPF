using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Services.Besancon;
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
        [NotifyPropertyChangedFor(nameof(EstEnSelectionScanMode))]
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
        private readonly SelectionScanModeViewModel _selectionScanModeViewModel = new();
        private readonly SelectionUtilisateurViewModel _selectionUtilisateurViewModel = new();

        public bool EstEnModeAdmin => VueActuelle is AdminViewModel;
        public bool EstEnSelectionPoste => VueActuelle is SelectionPosteViewModel;
        public bool EstEnSelectionScanMode => VueActuelle is SelectionScanModeViewModel;
        public bool EstEnSelectionUtilisateur => VueActuelle is SelectionUtilisateurViewModel;

        /// <summary>
        /// Faux tant que l'utilisateur n'a pas choisi son identité + son poste : la barre
        /// de navigation et la barre de statut sont alors masquées pour éviter qu'on
        /// contourne ces étapes en cliquant sur « Accueil ».
        /// </summary>
        public bool NavigationActive => !EstEnSelectionPoste && !EstEnSelectionScanMode && !EstEnSelectionUtilisateur;

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
                : "⚠ Vous n'utilisez pas la dernière configuration.\n"
                  + "Cliquez pour l'actualiser.\n\n"
                  + "Changements de configuration récents :\n"
                  + string.Join("\n", _changementsAdmin
                        .Skip(System.Math.Max(0, _changementsAdmin.Count - 8))
                        .Select(e => $"• {e.Horodatage:HH:mm} — {e.ActionLisible}"
                                   + (string.IsNullOrWhiteSpace(e.Utilisateur) ? "" : $" (par {e.Utilisateur})")));

        /// <summary>
        /// Pop-up rubidium sur CE poste (l'événement est déjà marshalé sur le thread UI par
        /// <see cref="BesanconScheduler"/>). Information (nouvelle valeur) ou avertissement (sous
        /// la limite 1e-13).
        /// </summary>
        private void OnPopupRubidiumDemandee(string titre, string message, bool avertissement)
        {
            MessageBox.Show(message, titre, MessageBoxButton.OK,
                avertissement ? MessageBoxImage.Warning : MessageBoxImage.Information);
        }

        private void OnChangementsAdminRecus(IReadOnlyList<EntreeJournalAdmin> nouveaux)
        {
            _changementsAdmin.AddRange(nouveaux);
            if (_changementsAdmin.Count > 100)
                _changementsAdmin.RemoveRange(0, _changementsAdmin.Count - 100);

            NbChangementsAdmin = _changementsAdmin.Count;
            AChangementsAdminEnAttente = true;
            OnPropertyChanged(nameof(ResumeChangementsAdmin));
        }

        /// <summary>Vrai quand un rafraîchissement a été demandé pendant une mesure et
        /// reste en attente de la fin de celle-ci (évite un double-abonnement).</summary>
        private bool _rafraichissementDiffereArme;

        /// <summary>
        /// Clic sur le triangle : affiche le détail des changements et propose de les
        /// charger tout de suite. « Oui » → relit les fichiers de configuration et applique
        /// les modifications ; si une mesure est en cours, on attend sa fin avant de relire.
        /// « Non » → l'indicateur reste affiché comme rappel ; les réglages seront pris au
        /// prochain démarrage.
        /// </summary>
        [RelayCommand]
        private async Task AcquitterChangementsAdmin()
        {
            if (_changementsAdmin.Count == 0)
            {
                AChangementsAdminEnAttente = false;
                NbChangementsAdmin = 0;
                return;
            }

            string liste = string.Join("\n", _changementsAdmin.Select(e =>
                $"• {e.Horodatage:dd/MM HH:mm} — {e.ActionLisible}"
              + (string.IsNullOrWhiteSpace(e.Detail) ? "" : $" : {e.Detail}")
              + (string.IsNullOrWhiteSpace(e.Utilisateur) ? "" : $"  (par {e.Utilisateur})")));

            var choix = MessageBox.Show(
                "Des changements de configuration ont été appliqués depuis un autre poste :\n\n" + liste
              + "\n\nVoulez-vous charger maintenant la dernière configuration ?\n\n"
              + "• Oui : relit les fichiers et applique les modifications (si une mesure est en\n"
              + "  cours, l'actualisation est différée jusqu'à la fin de la mesure).\n"
              + "• Non : un rappel reste affiché ; les nouveaux réglages seront pris au prochain démarrage.",
                "Changements administrateur", MessageBoxButton.YesNo, MessageBoxImage.Information);

            if (choix != MessageBoxResult.Yes)
                return; // « Plus tard » : on garde l'indicateur ⚠ comme rappel persistant.

            // Une mesure est en cours : on ne relit pas le fichier maintenant (on ne veut
            // pas changer les réglages au milieu d'une acquisition). On arme un
            // rafraîchissement différé qui s'exécutera dès la fin de la mesure.
            if (_accueilViewModel.MesureEnCours)
            {
                ArmerRafraichissementDiffere();
                MessageBox.Show(
                    "Une mesure est en cours.\n\nLa configuration sera actualisée automatiquement "
                  + "dès la fin de la mesure (le rappel ⚠ reste affiché jusque-là).",
                    "Actualisation différée", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            await ExecuterRafraichissementAsync();
        }

        /// <summary>S'abonne (une seule fois) à la fin de la mesure en cours pour relire la
        /// configuration dès qu'elle se termine.</summary>
        private void ArmerRafraichissementDiffere()
        {
            if (_rafraichissementDiffereArme) return;
            _rafraichissementDiffereArme = true;
            _accueilViewModel.PropertyChanged += OnAccueilPropertyChangedPourRafraichissement;
        }

        private void OnAccueilPropertyChangedPourRafraichissement(object? sender,
            System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName != nameof(AccueilViewModel.MesureEnCours)) return;
            if (_accueilViewModel.MesureEnCours) return; // mesure toujours en cours

            // Fin de mesure : on se désabonne et on applique. La fin de mesure peut survenir
            // sur un thread de fond → on marshale sur le thread UI avant de toucher aux
            // bindings (RubidiumActifChange, catalogue…).
            _accueilViewModel.PropertyChanged -= OnAccueilPropertyChangedPourRafraichissement;
            _rafraichissementDiffereArme = false;

            var dispatcher = System.Windows.Application.Current?.Dispatcher;
            if (dispatcher != null && !dispatcher.CheckAccess())
                _ = dispatcher.InvokeAsync(async () => await ExecuterRafraichissementAsync());
            else
                _ = ExecuterRafraichissementAsync();
        }

        /// <summary>Relit les fichiers de configuration et applique les changements, puis
        /// efface l'indicateur. Partagé entre l'actualisation immédiate et différée.</summary>
        private async Task ExecuterRafraichissementAsync()
        {
            try
            {
                await RafraichirConfigurationService.RafraichirAsync();

                _changementsAdmin.Clear();
                NbChangementsAdmin = 0;
                AChangementsAdminEnAttente = false;
                OnPropertyChanged(nameof(ResumeChangementsAdmin));

                ToastNotification.Afficher("Configuration à jour",
                    "Les derniers réglages administrateur ont été chargés.");
            }
            catch (System.Exception ex)
            {
                // Échec (ex. partage serveur injoignable) : on NE touche pas à l'indicateur
                // pour que l'utilisateur puisse réessayer plus tard.
                MessageBox.Show(
                    "Le rechargement de la configuration a échoué :\n" + ex.Message
                  + "\n\nLes réglages seront pris en compte au prochain démarrage.",
                    "Actualisation impossible", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        public MainViewModel()
        {
            NotificationsAdminWatcher.ChangementsRecus += OnChangementsAdminRecus;

            // Pop-up directe sur CE poste quand la valeur du rubidium est mise à jour (écart hebdo
            // Besançon) ou qu'elle passe sous la limite 1e-13. Le watcher ⚠ ne couvrant que les
            // AUTRES postes, cet événement assure l'affichage sur le poste qui exécute la tâche.
            BesanconScheduler.PopupRubidiumDemandee += OnPopupRubidiumDemandee;

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
                _ = Metrologo.Services.Journal.Journal.DefinirPosteAsync(choixBaie ? "Baie" : "Paillasse");

                if (choixBaie)
                {
                    // Baie : étape supplémentaire pour choisir Scan GPIB ou Adresses fixes.
                    VueActuelle = _selectionScanModeViewModel;
                }
                else
                {
                    // Paillasse : scan automatique, pas d'adresses fixes — on va direct à l'accueil.
                    EtatApplication.ModeAdressesFixes = false;
                    VueActuelle = _accueilViewModel;
                }
            };

            // Baie : choix Scan GPIB (true) ou Adresses fixes (false) → puis accueil.
            _selectionScanModeViewModel.OnModeChoisi = (estScan) =>
            {
                EtatApplication.ModeAdressesFixes = !estScan;
                VueActuelle = _accueilViewModel;
            };
            _selectionScanModeViewModel.OnRetour = () => VueActuelle = _selectionPosteViewModel;

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
