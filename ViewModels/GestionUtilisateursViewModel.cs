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
    /// <summary>
    /// CRUD local des utilisateurs (lecture / création / renommage / suppression / rôle)
    /// + gestion des mots de passe d'admin. Stockage 100 % local via
    /// <see cref="ComptesLocauxService"/>.
    ///
    /// Restrictions :
    ///   • Tout admin peut ajouter, renommer, supprimer un utilisateur lambda.
    ///   • Seul un super-administrateur peut modifier les rôles et réinitialiser
    ///     les mots de passe d'autres admins.
    ///   • Tout admin connecté peut changer son propre mot de passe.
    /// </summary>
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

        /// <summary>Vrai si l'admin authentifié dans la session courante est SuperAdmin.</summary>
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

        /// <summary>
        /// Change le rôle. Si la transition crée un nouvel admin (promotion depuis
        /// Utilisateur), le service génère un mot de passe et on l'affiche une seule
        /// fois pour qu'on puisse le communiquer.
        /// </summary>
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

        /// <summary>
        /// Réinitialise le mot de passe du compte admin sélectionné. Réservé au
        /// SuperAdministrateur. Le nouveau mot de passe est affiché une seule fois.
        /// </summary>
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

        /// <summary>
        /// Permet à l'admin connecté de changer son propre mot de passe. Accessible
        /// quel que soit le rôle (Admin ou SuperAdmin).
        /// </summary>
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
