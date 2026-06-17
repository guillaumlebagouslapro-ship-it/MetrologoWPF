using Metrologo.Models;
using Metrologo.Services;
using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    /// <summary>
    /// Laisse l'admin connecté (<see cref="EtatApplication.AdminConnecte"/>) changer son
    /// propre mot de passe. Comme il vient déjà de s'authentifier (login + mdp) pour entrer
    /// dans la zone admin, inutile de lui redemander l'ancien mot de passe — ça permet
    /// d'ailleurs de remplacer un mdp auto-généré qu'il n'a jamais retenu. Il tape donc
    /// directement le nouveau deux fois, et on enregistre via
    /// <see cref="ComptesLocauxService.DefinirMotDePasse"/>.
    /// </summary>
    public partial class ChangerMdpAdminWindow : FluentWindow
    {
        public ChangerMdpAdminWindow()
        {
            InitializeComponent();
            Loaded += (_, _) => PbNouveau.Focus();
        }

        private void OnValider(object sender, RoutedEventArgs e)
        {
            var admin = EtatApplication.AdminConnecte;
            if (admin == null)
            {
                AfficherErreur("Aucun admin connecté — ré-ouvre la zone admin.");
                return;
            }

            string nouveau = PbNouveau.Password ?? string.Empty;
            string confirmer = PbConfirmer.Password ?? string.Empty;

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
            // Le hash gardé en mémoire dans EtatApplication.AdminConnecte est désormais
            // périmé : on le réaligne tout de suite pour que les vérifs suivantes, dans
            // la même session admin, retombent bien sur le nouveau hash.
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
