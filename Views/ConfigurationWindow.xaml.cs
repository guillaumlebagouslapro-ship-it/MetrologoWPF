using System.Windows;
using System.Windows.Controls;
using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class ConfigurationWindow : FluentWindow
    {
        /// <summary>
        /// Garde-fou : quand le handler TextChanged retouche lui-même le contenu du
        /// TextBox, ça relance un TextChanged. Ce drapeau coupe la boucle pour ne pas
        /// tourner en rond.
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
            // On laisse le ViewModel faire son travail
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

            // Est-ce que l'utilisateur vient d'effacer (backspace, delete) ? Si oui,
            // on s'abstient de remettre le _, faute de quoi le champ deviendrait
            // impossible à éditer une fois les 2 premiers caractères tapés.
            bool suppression = false;
            foreach (var ch in e.Changes)
            {
                if (ch.RemovedLength > ch.AddedLength) { suppression = true; break; }
            }

            string texte = tb.Text ?? string.Empty;
            string nouveau = texte;

            // Première lettre en majuscule (D, E, F, G…) ; on ne touche à rien si elle l'est déjà
            if (nouveau.Length >= 1 && char.IsLetter(nouveau[0]) && !char.IsUpper(nouveau[0]))
            {
                nouveau = char.ToUpper(nouveau[0]) + (nouveau.Length > 1 ? nouveau.Substring(1) : string.Empty);
            }

            // Le séparateur _ va en 3e position :
            //   - si on vient de taper le 3e caractère, on l'accroche au bout
            //   - si l'utilisateur a effacé le _ mais qu'il reste du texte derrière, on le remet
            //     (sans ça, le format XX_NNNNN se casserait sans prévenir)
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
                tb.CaretIndex = nouveau.Length;   // on remet le curseur à la fin
                _miseAJourNumFI = false;
            }

            // On redemande l'évaluation de l'étape 1 (FI valide), c'est elle qui décide
            // de l'affichage des sections suivantes.
            (DataContext as ConfigurationViewModel)?.NotifierEtapes();
        }
    }
}