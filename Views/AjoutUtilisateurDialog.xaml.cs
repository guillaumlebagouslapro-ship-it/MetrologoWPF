using System.Windows;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class AjoutUtilisateurDialog : FluentWindow
    {
        public string NomSaisi => TbNom.Text.Trim();
        public string PrenomSaisi => TbPrenom.Text.Trim();

        public AjoutUtilisateurDialog()
        {
            InitializeComponent();
            Loaded += (_, _) => TbPrenom.Focus();
        }

        /// <summary>
        /// Le même dialogue sert aussi à renommer : on pré-remplit les champs et on adapte
        /// le titre, le libellé du bouton et le texte d'intro. Du coup le ViewModel appelle
        /// Renommer() plutôt qu'Ajouter().
        /// </summary>
        public void ConfigurerPourRenommage(string nomActuel, string prenomActuel)
        {
            Title = "Modifier l'utilisateur";
            TbDescription.Text = "Modifier le nom et le prénom. Le login reste inchangé.";
            BtnValider.Content = "Enregistrer";
            TbPrenom.Text = prenomActuel;
            TbNom.Text = nomActuel;
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
