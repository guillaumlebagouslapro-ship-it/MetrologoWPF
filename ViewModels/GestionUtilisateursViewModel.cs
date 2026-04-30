using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
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
    public partial class GestionUtilisateursViewModel : ObservableObject
    {
        private readonly IUtilisateursService _service;

        public ObservableCollection<Utilisateur> Utilisateurs { get; } = new();

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(SupprimerCommand))]
        [NotifyCanExecuteChangedFor(nameof(ReinitialiserMotDePasseCommand))]
        private Utilisateur? _selection;

        [ObservableProperty] private string _statut = string.Empty;
        [ObservableProperty] private bool _enChargement;

        public GestionUtilisateursViewModel(IUtilisateursService service)
        {
            _service = service;
            _ = ChargerAsync();
        }

        [RelayCommand]
        private async Task RafraichirAsync() => await ChargerAsync();

        private async Task ChargerAsync()
        {
            try
            {
                EnChargement = true;
                Statut = "Chargement…";
                Utilisateurs.Clear();
                var liste = await _service.ListerAsync();
                foreach (var u in liste) Utilisateurs.Add(u);
                Statut = $"{Utilisateurs.Count} utilisateur(s).";
            }
            catch (Exception ex)
            {
                Statut = $"Erreur : {ex.Message}";
                JournalLog.Warn(CategorieLog.Administration, "GESTION_USERS_LOAD_ERR", ex.Message);
            }
            finally { EnChargement = false; }
        }

        [RelayCommand]
        private async Task AjouterAsync()
        {
            var dlg = new AjoutUtilisateurDialog { Owner = Application.Current.MainWindow };
            if (dlg.ShowDialog() != true) return;

            try
            {
                var (utilisateur, mdpClair) = await _service.CreerAsync(
                    dlg.NomSaisi, dlg.PrenomSaisi, dlg.RoleSelectionne);

                Utilisateurs.Add(utilisateur);
                Selection = utilisateur;

                // Affichage du mot de passe une seule fois (à communiquer à l'utilisateur).
                var info = new InfoCompteCreeDialog(utilisateur.Login, mdpClair) { Owner = Application.Current.MainWindow };
                info.ShowDialog();

                Statut = $"Compte {utilisateur.Login} créé.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Création échouée : {ex.Message}", "Erreur",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        [RelayCommand(CanExecute = nameof(PeutSupprimer))]
        private async Task SupprimerAsync()
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
                bool ok = await _service.SupprimerAsync(Selection.Id);
                if (ok)
                {
                    Utilisateurs.Remove(Selection);
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

        private bool PeutSupprimer() => Selection != null;

        [RelayCommand(CanExecute = nameof(PeutReinitialiser))]
        private async Task ReinitialiserMotDePasseAsync()
        {
            if (Selection == null) return;

            var conf = MessageBox.Show(
                $"Réinitialiser le mot de passe de {Selection.NomComplet} ({Selection.Login}) ?\n\n" +
                "L'ancien mot de passe sera immédiatement invalidé. Le nouveau mot de passe sera affiché " +
                "une seule fois — à communiquer à l'utilisateur.",
                "Confirmer la réinitialisation",
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (conf != MessageBoxResult.Yes) return;

            try
            {
                string nouveauMdp = await _service.ReinitialiserMotDePasseAsync(Selection.Id);

                var info = new InfoCompteCreeDialog(Selection.Login, nouveauMdp,
                    titre: "Mot de passe réinitialisé") { Owner = Application.Current.MainWindow };
                info.ShowDialog();

                Statut = $"Mot de passe réinitialisé pour {Selection.Login}.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Réinitialisation échouée : {ex.Message}", "Erreur",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool PeutReinitialiser() => Selection != null;
    }
}
