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
        /// Auto-scroll vers la dernière ligne du journal "Informations générales" à chaque
        /// ajout — évite à l'utilisateur de devoir faire défiler manuellement pour voir les
        /// nouvelles entrées pendant une mesure qui s'allonge.
        /// </summary>
        private void LogTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox tb) tb.ScrollToEnd();
        }
    }
}
