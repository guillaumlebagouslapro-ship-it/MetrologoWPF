using Metrologo.Models;
using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class ChangerRoleDialog : FluentWindow
    {
        public RoleUtilisateur RoleChoisi { get; private set; }

        public ChangerRoleDialog(Utilisateur utilisateur)
        {
            InitializeComponent();

            TbNomComplet.Text = utilisateur.NomComplet;
            TbLogin.Text = utilisateur.Login;
            TbInitiales.Text = Initiales(utilisateur);

            switch (utilisateur.Role)
            {
                case RoleUtilisateur.SuperAdministrateur: RbSuperAdmin.IsChecked = true; break;
                case RoleUtilisateur.Administrateur:     RbAdministrateur.IsChecked = true; break;
                default:                                  RbUtilisateur.IsChecked = true; break;
            }

            RoleChoisi = utilisateur.Role;
        }

        private static string Initiales(Utilisateur u)
        {
            char p = !string.IsNullOrEmpty(u.Prenom) ? char.ToUpper(u.Prenom[0]) : '?';
            char n = !string.IsNullOrEmpty(u.Nom) ? char.ToUpper(u.Nom[0]) : '?';
            return $"{p}{n}";
        }

        private void OnValider(object sender, RoutedEventArgs e)
        {
            if (RbSuperAdmin.IsChecked == true) RoleChoisi = RoleUtilisateur.SuperAdministrateur;
            else if (RbAdministrateur.IsChecked == true) RoleChoisi = RoleUtilisateur.Administrateur;
            else RoleChoisi = RoleUtilisateur.Utilisateur;

            DialogResult = true;
        }

        private void OnAnnuler(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
