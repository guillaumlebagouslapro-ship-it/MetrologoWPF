using System.Windows;
using System.Windows.Controls;
using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class ConfigurationWindow : FluentWindow
    {
        /// <summary>
        /// Évite la récursion infinie quand le handler TextChanged modifie lui-même
        /// le contenu du TextBox (ce qui re-déclenche TextChanged).
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
            // On appelle la logique du ViewModel
            (DataContext as ConfigurationViewModel)?.OnTypeMesureChanged();
        }

        /// <summary>
        /// Auto-format de la saisie du N° FI au format <c>XX_NNNNN</c> (8 caractères) :
        /// <list type="bullet">
        ///   <item>1ère lettre forcée en MAJUSCULE (les FI sont stockées en majuscule en BDD).</item>
        ///   <item>Insère automatiquement <c>_</c> après les 2 premiers caractères saisis.</item>
        ///   <item>Ne se ré-insère PAS si l'utilisateur efface (backspace OK pour corriger).</item>
        /// </list>
        /// </summary>
        private void OnNumFITextChanged(object sender, TextChangedEventArgs e)
        {
            if (_miseAJourNumFI) return;
            if (sender is not System.Windows.Controls.TextBox tb) return;

            // Détecte si l'utilisateur vient de SUPPRIMER (backspace, delete) :
            // dans ce cas on ne ré-insère pas le _, sinon il serait impossible
            // d'éditer le champ après avoir tapé 2 caractères.
            bool suppression = false;
            foreach (var ch in e.Changes)
            {
                if (ch.RemovedLength > ch.AddedLength) { suppression = true; break; }
            }

            string texte = tb.Text ?? string.Empty;
            string nouveau = texte;

            // 1ère lettre en majuscule (D, E, F, G…) — pas d'effet si déjà majuscule
            if (nouveau.Length >= 1 && char.IsLetter(nouveau[0]) && !char.IsUpper(nouveau[0]))
            {
                nouveau = char.ToUpper(nouveau[0]) + (nouveau.Length > 1 ? nouveau.Substring(1) : string.Empty);
            }

            // Gestion du séparateur _ en position 3 :
            //   - on vient de taper le 3e caractère → l'ajouter en bout
            //   - l'utilisateur a effacé le _ alors qu'il y a du texte derrière → le réinsérer
            //     (sinon le format XX_NNNNN serait cassé silencieusement)
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
                tb.CaretIndex = nouveau.Length;   // place le curseur en fin
                _miseAJourNumFI = false;
            }

            // Re-évalue l'étape 1 (FI valide) qui pilote l'affichage des sections suivantes.
            (DataContext as ConfigurationViewModel)?.NotifierEtapes();
        }
    }
}