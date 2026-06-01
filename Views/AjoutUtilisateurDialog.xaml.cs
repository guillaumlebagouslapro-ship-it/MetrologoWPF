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
        /// Sert aussi pour le renommage : pré-remplit les champs et change le titre +
        /// libellé du bouton + texte d'intro. Le ViewModel appelle ensuite Renommer()
        /// au lieu d'Ajouter().
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
