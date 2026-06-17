using System.Windows.Controls;

namespace Metrologo.Views
{
    public partial class AcceuilView : UserControl
    {
        public AcceuilView()
        {
            InitializeComponent();
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
