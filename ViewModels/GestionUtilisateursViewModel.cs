using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Services.Journal;
using Metrologo.Views;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.ViewModels
{
    /// <summary>CRUD des utilisateurs + mots de passe admin, stockage local via ComptesLocauxService.
    /// Un admin gère les utilisateurs lambda et son propre mot de passe ; seul un SuperAdmin
    /// touche aux rôles et aux mots de passe des autres admins.</summary>
    public partial class GestionUtilisateursViewModel : ObservableObject
    {
        public ObservableCollection<Utilisateur> Utilisateurs { get; } = new();

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(SupprimerCommand))]
        [NotifyCanExecuteChangedFor(nameof(RenommerCommand))]
        [NotifyCanExecuteChangedFor(nameof(ChangerRoleCommand))]
        [NotifyCanExecuteChangedFor(nameof(ReinitialiserMotDePasseCommand))]
        private Utilisateur? _selection;

        [ObservableProperty] private string _statut = string.Empty;

        // vrai si l'admin de la session courante est SuperAdmin
        public bool EstSuperAdmin => EtatApplication.EstSuperAdmin;

        public GestionUtilisateursViewModel()
        {
            Charger();
        }

        [RelayCommand]
        private void Rafraichir() => Charger();

        private void Charger()
        {
            try
            {
                // on relit le JSON à jour : Rafraîchir voit ainsi les comptes créés ou
                // modifiés depuis un autre poste sans redémarrer l'appli
                Preferences.InvaliderCacheUtilisateurs();

                Utilisateurs.Clear();
                foreach (var u in ComptesLocauxService.Lister()) Utilisateurs.Add(u);
                Statut = $"{Utilisateurs.Count} utilisateur(s).";
                OnPropertyChanged(nameof(EstSuperAdmin));
            }
            catch (Exception ex)
            {
                Statut = $"Erreur : {ex.Message}";
                JournalLog.Warn(CategorieLog.Administration, "GESTION_USERS_LOAD_ERR", ex.Message);
            }
        }

        [RelayCommand]
        private void Ajouter()
        {
            var dlg = new AjoutUtilisateurDialog { Owner = Application.Current.MainWindow };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var utilisateur = ComptesLocauxService.Ajouter(dlg.NomSaisi, dlg.PrenomSaisi);
                Charger();
                Selection = Utilisateurs.FirstOrDefault(u => u.Id == utilisateur.Id);
                Statut = $"Compte {utilisateur.Login} créé.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Création échouée : {ex.Message}", "Erreur",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        [RelayCommand(CanExecute = nameof(PeutModifier))]
        private void Renommer()
        {
            if (Selection == null) return;

            var dlg = new AjoutUtilisateurDialog { Owner = Application.Current.MainWindow };
            dlg.ConfigurerPourRenommage(Selection.Nom, Selection.Prenom);
            if (dlg.ShowDialog() != true) return;

            try
            {
                if (ComptesLocauxService.Renommer(Selection.Id, dlg.NomSaisi, dlg.PrenomSaisi))
                {
                    int idCourant = Selection.Id;
                    Charger();
                    Selection = Utilisateurs.FirstOrDefault(u => u.Id == idCourant);
                    Statut = "Utilisateur modifié.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Modification échouée : {ex.Message}", "Erreur",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        [RelayCommand(CanExecute = nameof(PeutModifier))]
        private void Supprimer()
        {
            if (Selection == null) return;

            var conf = MessageBox.Show(
                $"Supprimer l'utilisateur {Selection.NomComplet} ({Selection.Login}) ?\n\n" +
                "Cette action est irréversible.",
                "Confirmer la suppression",
                MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (conf != MessageBoxResult.Yes) return;

            try
            {
                if (ComptesLocauxService.Supprimer(Selection.Id))
                {
                    Utilisateurs.Remove(Selection);
                    Selection = null;
                    Statut = "Utilisateur supprimé.";
                }
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message, "Suppression refusée",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Suppression échouée : {ex.Message}", "Erreur",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // change le rôle ; sur une promotion en admin le service génère un mot de passe,
        // affiché une seule fois pour le communiquer à l'intéressé
        [RelayCommand(CanExecute = nameof(PeutChangerRole))]
        private void ChangerRole()
        {
            if (Selection == null) return;

            var dlg = new ChangerRoleDialog(Selection) { Owner = Application.Current.MainWindow };
            if (dlg.ShowDialog() != true) return;
            if (dlg.RoleChoisi == Selection.Role) return;

            try
            {
                string? mdpGenere = ComptesLocauxService.ChangerRole(Selection.Id, dlg.RoleChoisi);

                int idCourant = Selection.Id;
                Charger();
                Selection = Utilisateurs.FirstOrDefault(u => u.Id == idCourant);

                if (mdpGenere != null && Selection != null)
                {
                    var info = new InfoMdpGenereDialog(Selection.Login, mdpGenere,
                        titre: "Compte promu administrateur")
                    { Owner = Application.Current.MainWindow };
                    info.ShowDialog();
                }

                Statut = $"Rôle de {Selection?.Login ?? "?"} : {dlg.RoleChoisi}.";
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message, "Action refusée",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Modification échouée : {ex.Message}", "Erreur",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // réinitialise le mot de passe de l'admin sélectionné (SuperAdmin uniquement),
        // le nouveau n'est affiché qu'une seule fois
        [RelayCommand(CanExecute = nameof(PeutReinitialiserMotDePasse))]
        private void ReinitialiserMotDePasse()
        {
            if (Selection == null) return;

            var conf = MessageBox.Show(
                $"Réinitialiser le mot de passe de {Selection.NomComplet} ({Selection.Login}) ?\n\n" +
                "L'ancien mot de passe sera immédiatement invalidé. Le nouveau sera affiché une seule fois.",
                "Confirmer la réinitialisation",
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (conf != MessageBoxResult.Yes) return;

            try
            {
                string nouveauMdp = ComptesLocauxService.ReinitialiserMotDePasse(Selection.Id);
                var info = new InfoMdpGenereDialog(Selection.Login, nouveauMdp,
                    titre: "Mot de passe réinitialisé")
                { Owner = Application.Current.MainWindow };
                info.ShowDialog();
                Statut = $"Mot de passe réinitialisé pour {Selection.Login}.";
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message, "Action refusée",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Réinitialisation échouée : {ex.Message}", "Erreur",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // l'admin connecté change son propre mot de passe (Admin comme SuperAdmin)
        [RelayCommand]
        private void ChangerMonMotDePasse()
        {
            var dlg = new ChangerMdpAdminWindow { Owner = Application.Current.MainWindow };
            if (dlg.ShowDialog() == true)
            {
                Statut = "Mot de passe modifié.";
            }
        }

        private bool PeutModifier() => Selection != null;
        private bool PeutChangerRole() => Selection != null && EstSuperAdmin;
        private bool PeutReinitialiserMotDePasse() =>
            Selection != null
            && Selection.Role != RoleUtilisateur.Utilisateur
            && EstSuperAdmin;
    }
}
