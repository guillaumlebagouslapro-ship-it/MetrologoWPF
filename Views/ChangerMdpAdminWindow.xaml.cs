using Metrologo.Models;
using Metrologo.Services;
using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    /// <summary>
    /// Permet à l'admin actuellement connecté (<see cref="EtatApplication.AdminConnecte"/>)
    /// de changer son propre mot de passe. On vérifie l'ancien, on demande deux fois le
    /// nouveau, puis on enregistre via <see cref="ComptesLocauxService.DefinirMotDePasse"/>.
    /// </summary>
    public partial class ChangerMdpAdminWindow : FluentWindow
    {
        public ChangerMdpAdminWindow()
        {
            InitializeComponent();
            Loaded += (_, _) => PbActuel.Focus();
        }

        private void OnValider(object sender, RoutedEventArgs e)
        {
            var admin = EtatApplication.AdminConnecte;
            if (admin == null)
            {
                AfficherErreur("Aucun admin connecté — ré-ouvre la zone admin.");
                return;
            }

            string actuel = PbActuel.Password ?? string.Empty;
            string nouveau = PbNouveau.Password ?? string.Empty;
            string confirmer = PbConfirmer.Password ?? string.Empty;

            // Vérifie l'ancien mdp contre le hash du compte connecté.
            if (admin.PasswordHash == null || !PasswordHasher.VerifyPassword(actuel, admin.PasswordHash))
            {
                AfficherErreur("Mot de passe actuel incorrect.");
                PbActuel.Clear();
                PbActuel.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(nouveau))
            {
                AfficherErreur("Le nouveau mot de passe est requis.");
                PbNouveau.Focus();
                return;
            }

            if (nouveau.Length < 4)
            {
                AfficherErreur("Le nouveau mot de passe doit faire au moins 4 caractères.");
                PbNouveau.Focus();
                return;
            }

            if (nouveau != confirmer)
            {
                AfficherErreur("Les deux saisies du nouveau mot de passe ne correspondent pas.");
                PbConfirmer.Clear();
                PbConfirmer.Focus();
                return;
            }

            ComptesLocauxService.DefinirMotDePasse(admin.Id, nouveau);
            // Le hash en mémoire de EtatApplication.AdminConnecte n'est plus à jour.
            // On le rafraîchit pour que les vérifications ultérieures dans la même
            // session admin matchent le nouveau hash.
            admin.PasswordHash = PasswordHasher.HashPassword(nouveau);

            DialogResult = true;
            Close();
        }

        private void OnAnnuler(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void AfficherErreur(string message)
        {
            TbErreur.Text = message;
            BdErreur.Visibility = Visibility.Visible;
        }
    }
}
