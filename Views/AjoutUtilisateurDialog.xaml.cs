using System.Windows;
using Metrologo.Models;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class AjoutUtilisateurDialog : FluentWindow
    {
        public string NomSaisi => TbNom.Text.Trim();
        public string PrenomSaisi => TbPrenom.Text.Trim();
        public RoleUtilisateur RoleSelectionne =>
            CbRole.SelectedIndex == 1 ? RoleUtilisateur.Administrateur : RoleUtilisateur.Utilisateur;

        public AjoutUtilisateurDialog()
        {
            InitializeComponent();
            Loaded += (_, _) => TbPrenom.Focus();
        }

        private void OnValider(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(PrenomSaisi))
            {
                System.Windows.MessageBox.Show("Le prénom est requis.", "Champ manquant",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                TbPrenom.Focus();
                return;
            }
            if (string.IsNullOrWhiteSpace(NomSaisi))
            {
                System.Windows.MessageBox.Show("Le nom est requis.", "Champ manquant",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                TbNom.Focus();
                return;
            }
            DialogResult = true;
        }

        private void OnAnnuler(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
