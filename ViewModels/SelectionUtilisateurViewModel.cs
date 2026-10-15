using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services;
using System.Collections.ObjectModel;
using System.Globalization;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// C'est le tout premier écran au démarrage : un menu déroulant qui présente les comptes
    /// locaux. L'utilisateur indique simplement qui se connecte (aucun mot de passe à ce
    /// stade), puis on enchaîne sur le choix Baie / Paillasse.
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

        /// <summary>Date du jour à la française (ex. « mardi 22 mai 2026 »), rafraîchie par le timer
        /// car l'écran peut rester ouvert des heures.</summary>
        public string DateActuelle =>
            CultureInfo.GetCultureInfo("fr-FR").TextInfo.ToTitleCase(
                DateTime.Now.ToString("dddd d MMMM yyyy", CultureInfo.GetCultureInfo("fr-FR")));

        /// <summary>Heure courante au format HH:mm, elle aussi rafraîchie en continu.</summary>
        public string HeureActuelle => DateTime.Now.ToString("HH:mm");

        // Timer UI : renotifie DateActuelle/HeureActuelle chaque seconde (évite un décalage
        // visible à l'ouverture de l'écran, bien que l'affichage ne change qu'à la minute).
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
        /// Recharge la liste à partir du service local. À déclencher au retour de la fenêtre
        /// Gestion des utilisateurs pour que les comptes fraîchement créés apparaissent.
        /// </summary>
        public void Recharger()
        {
            // On invalide le cache mémoire pour relire le JSON à jour. C'est ce qui permet de
            // voir un compte tout juste créé, même depuis un autre poste, sans relancer l'appli.
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
