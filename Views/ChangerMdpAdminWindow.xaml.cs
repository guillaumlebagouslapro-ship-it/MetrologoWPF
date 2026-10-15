using Metrologo.Models;
using Metrologo.Services;
using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    /// <summary>
    /// Changement du mot de passe de l'admin connecté (<see cref="EtatApplication.AdminConnecte"/>).
    /// L'ancien mot de passe n'est pas redemandé (déjà authentifié pour entrer en zone admin ;
    /// permet aussi de remplacer un mdp auto-généré jamais retenu). Enregistrement via
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
            // Le hash en mémoire dans EtatApplication.AdminConnecte est périmé : on le réaligne
            // pour les vérifications suivantes de la même session admin.
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
