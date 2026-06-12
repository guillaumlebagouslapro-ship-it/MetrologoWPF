using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services;
using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// Premier écran affiché au démarrage : un menu déroulant qui liste les comptes
    /// locaux. L'utilisateur choisit qui se connecte (pas de mot de passe à ce stade),
    /// puis on enchaîne sur la sélection Baie / Paillasse.
    /// </summary>
    public partial class SelectionUtilisateurViewModel : ObservableObject
    {
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(ListeVide))]
        [NotifyPropertyChangedFor(nameof(NombreComptesTexte))]
        [NotifyCanExecuteChangedFor(nameof(ContinuerCommand))]
        private ObservableCollection<Utilisateur> _utilisateurs = new();

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(ContinuerCommand))]
        private Utilisateur? _utilisateurSelectionne;

        public bool ListeVide => Utilisateurs.Count == 0;

        /// <summary>Date du jour formatée en français (ex. « mardi 22 mai 2026 »).
        /// Rafraîchie en continu par le timer (l'écran de lancement peut rester
        /// affiché des heures — la date/heure doit suivre).</summary>
        public string DateActuelle =>
            CultureInfo.GetCultureInfo("fr-FR").TextInfo.ToTitleCase(
                DateTime.Now.ToString("dddd d MMMM yyyy", CultureInfo.GetCultureInfo("fr-FR")));

        /// <summary>Heure du jour formatée HH:mm, rafraîchie en continu.</summary>
        public string HeureActuelle => DateTime.Now.ToString("HH:mm");

        // Timer UI qui notifie DateActuelle/HeureActuelle chaque seconde — l'affichage
        // ne change visuellement qu'au changement de minute, mais le tick fin évite un
        // décalage perceptible juste après l'ouverture de l'écran.
        private readonly System.Windows.Threading.DispatcherTimer _horloge;

        public string NombreComptesTexte => Utilisateurs.Count switch
        {
            0 => "aucun compte",
            1 => "1 compte enregistré",
            _ => $"{Utilisateurs.Count} comptes enregistrés"
        };

        public string VersionApp => "v2026";

        public Action<Utilisateur>? OnUtilisateurChoisi { get; set; }

        public SelectionUtilisateurViewModel()
        {
            Recharger();

            _horloge = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _horloge.Tick += (_, _) =>
            {
                OnPropertyChanged(nameof(HeureActuelle));
                OnPropertyChanged(nameof(DateActuelle));
            };
            _horloge.Start();
        }

        /// <summary>
        /// Recharge la liste depuis le service local. À appeler après un retour de la
        /// fenêtre Gestion des utilisateurs si on veut refléter une création récente.
        /// </summary>
        public void Recharger()
        {
            // Relit le JSON à jour (cache mémoire invalidé) — indispensable pour voir un compte
            // tout juste créé, y compris depuis un autre poste, sans redémarrer l'application.
            Preferences.InvaliderCacheUtilisateurs();

            Utilisateurs.Clear();
            foreach (var u in ComptesLocauxService.Lister()) Utilisateurs.Add(u);
            UtilisateurSelectionne = Utilisateurs.FirstOrDefault();
            OnPropertyChanged(nameof(ListeVide));
        }

        [RelayCommand(CanExecute = nameof(PeutContinuer))]
        private void Continuer()
        {
            if (UtilisateurSelectionne == null) return;
            OnUtilisateurChoisi?.Invoke(UtilisateurSelectionne);
        }

        private bool PeutContinuer() => UtilisateurSelectionne != null;
    }
}
