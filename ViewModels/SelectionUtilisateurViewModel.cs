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

        /// <summary>Nom du poste (machine Windows) — affiché dans le panneau d'accueil.</summary>
        public string NomPoste => Environment.MachineName;

        /// <summary>Date du jour formatée en français (ex. « mardi 22 mai 2026 »).</summary>
        public string DateActuelle =>
            CultureInfo.GetCultureInfo("fr-FR").TextInfo.ToTitleCase(
                DateTime.Now.ToString("dddd d MMMM yyyy", CultureInfo.GetCultureInfo("fr-FR")));

        /// <summary>Heure du jour formatée HH:mm.</summary>
        public string HeureActuelle => DateTime.Now.ToString("HH:mm");

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
