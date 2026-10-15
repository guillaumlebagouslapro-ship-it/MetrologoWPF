using System.Windows.Controls;
using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class ConfigurationWindow : FluentWindow
    {
        /// <summary>
        /// Garde-fou : le handler TextChanged retouche le TextBox, ce qui le relance ; ce drapeau coupe la boucle.
        /// </summary>
        private bool _miseAJourNumFI;

        public ConfigurationWindow(ConfigurationViewModel vm)
        {
            InitializeComponent();
            DataContext = vm;

            vm.CloseAction = result =>
            {
                this.DialogResult = result;
            };
        }

        private void ComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            (DataContext as ConfigurationViewModel)?.RefreshAll();
        }

        private void TypeMesure_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            (DataContext as ConfigurationViewModel)?.OnTypeMesureChanged();
        }

        /// <summary>
        /// Met en forme la saisie du N° FI à la volée, au format <c>XX_NNNNN</c> (8 caractères) :
        /// <list type="bullet">
        ///   <item>la 1ère lettre passe en MAJUSCULE (c'est ainsi que les FI sont stockées en BDD).</item>
        ///   <item>le <c>_</c> s'insère tout seul après les 2 premiers caractères.</item>
        ///   <item>mais on ne le remet pas quand l'utilisateur efface, pour qu'il puisse corriger au backspace.</item>
        /// </list>
        /// </summary>
        private void OnNumFITextChanged(object sender, TextChangedEventArgs e)
        {
            if (_miseAJourNumFI) return;
            if (sender is not System.Windows.Controls.TextBox tb) return;

            // Suppression (backspace, delete) : on ne remet pas le _, sinon le champ
            // deviendrait impossible à éditer une fois les 2 premiers caractères tapés.
            bool suppression = false;
            foreach (var ch in e.Changes)
            {
                if (ch.RemovedLength > ch.AddedLength) { suppression = true; break; }
            }

            string texte = tb.Text ?? string.Empty;
            string nouveau = texte;

            // Première lettre en majuscule (D, E, F, G…)
            if (nouveau.Length >= 1 && char.IsLetter(nouveau[0]) && !char.IsUpper(nouveau[0]))
            {
                nouveau = char.ToUpper(nouveau[0]) + (nouveau.Length > 1 ? nouveau.Substring(1) : string.Empty);
            }

            // Le _ va en 3e position : ajouté après 2 caractères tapés, ou remis s'il a été
            // effacé alors que du texte le suit (sinon le format XX_NNNNN se casserait).
            if (!nouveau.Contains('_'))
            {
                if (nouveau.Length == 2 && !suppression)
                {
                    nouveau += "_";
                }
                else if (nouveau.Length >= 3)
                {
                    nouveau = nouveau.Substring(0, 2) + "_" + nouveau.Substring(2);
                }
            }

            if (nouveau != texte)
            {
                _miseAJourNumFI = true;
                tb.Text = nouveau;
                tb.CaretIndex = nouveau.Length;
                _miseAJourNumFI = false;
            }

            // Réévalue l'étape 1 (FI valide), qui commande l'affichage des sections suivantes.
            (DataContext as ConfigurationViewModel)?.NotifierEtapes();
        }
    }
}