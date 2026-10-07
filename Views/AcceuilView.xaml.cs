using System.Windows.Controls;

namespace Metrologo.Views
{
    public partial class AcceuilView : UserControl
    {
        public AcceuilView()
        {
            InitializeComponent();

            string? version = UpdateService.VersionInstallee();
            if (version != null)
                MisesAJourTexte.Text = "Version " + HistoriqueMaj.Affichage(version);
        }

        /// <summary>Accueil > Mises à jour : historique + retour à la version précédente.</summary>
        private void MisesAJour_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            new MisesAJourWindow { Owner = System.Windows.Window.GetWindow(this) }.ShowDialog();
        }

        /// <summary>
        /// On garde toujours la dernière ligne du journal "Informations générales" visible :
        /// à chaque nouvel ajout on descend automatiquement en bas. Comme ça, pendant une
        /// mesure qui s'étire, l'utilisateur n'a pas à faire défiler à la main pour suivre ce
        /// qui se passe.
        /// </summary>
        private void LogTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox tb) tb.ScrollToEnd();
        }
    }
}
