using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Metrologo.Services.Incertitude;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class AjoutTrioDialog : FluentWindow
    {
        // Liste canonique des temps de mesure standards proposés à l'admin.
        private static readonly double[] _tempsStandards =
        {
            0.01, 0.02, 0.05, 0.1, 0.2, 0.5, 1d, 2d, 5d, 10d, 20d, 50d, 100d
        };

        private readonly ModuleIncertitude? _module;

        public string Fonction { get; private set; } = "Freq";
        public double TempsDeMesure { get; private set; } = 1.0;

        /// <summary>1 pour une plage unique, 3 pour le trio basse/moyenne/haute (défaut).</summary>
        public int NombrePlages { get; private set; } = 3;

        public AjoutTrioDialog() : this(null) { }

        /// <summary>
        /// <paramref name="module"/> = module en cours d'édition. Si fourni, la
        /// ComboBox des temps est filtrée pour exclure les temps déjà saisis pour
        /// la fonction sélectionnée (évite les doublons).
        /// </summary>
        public AjoutTrioDialog(ModuleIncertitude? module)
        {
            InitializeComponent();
            _module = module;

            // Module sans temps de mesure : on masque tout le bloc et on fixe TempsDeMesure
            // à 0 dans OnValider — la sélection du temps ne fait pas sens (tachy/strobo).
            // Dans ce cas on pré-sélectionne aussi « 1 plage » (cas par défaut tachy).
            if (module != null && !module.UtiliseTempsDeMesure)
            {
                PanelTemps.Visibility = System.Windows.Visibility.Collapsed;
                Title = "Ajouter une plage de fréquence";
                Height = 320;
                RbUnePlage.IsChecked = true;
            }
            else
            {
                CbFonction.SelectionChanged += (_, _) => RafraichirTempsDisponibles();
                Loaded += (_, _) => RafraichirTempsDisponibles();
            }

            // Le libellé du bouton suit le choix de l'utilisateur en temps réel.
            RbUnePlage.Checked   += (_, _) => MajLibelleBouton();
            RbTroisPlages.Checked += (_, _) => MajLibelleBouton();
            Loaded += (_, _) => MajLibelleBouton();
        }

        private void MajLibelleBouton()
        {
            BtnCreer.Content = RbUnePlage.IsChecked == true
                ? "Créer la plage"
                : "Créer les 3 plages";
        }

        private void RafraichirTempsDisponibles()
        {
            string fnSel = (CbFonction.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Freq";

            // Temps déjà présents dans le tableau pour cette fonction (sur le module courant).
            HashSet<double> dejaPresents = _module == null
                ? new HashSet<double>()
                : new HashSet<double>(_module.Lignes
                    .Where(l => string.Equals(l.Fonction, fnSel, System.StringComparison.OrdinalIgnoreCase))
                    .Select(l => l.TempsDeMesure));

            // Garde le texte courant (pour ne pas écraser une saisie en cours).
            string textePrecedent = CbTemps.Text;

            CbTemps.Items.Clear();
            foreach (double t in _tempsStandards)
            {
                if (dejaPresents.Contains(t)) continue; // exclu : déjà présent
                CbTemps.Items.Add(new ComboBoxItem
                {
                    Content = t.ToString(CultureInfo.InvariantCulture)
                });
            }

            // Sélectionne le 1er restant si la liste n'est pas vide.
            if (CbTemps.Items.Count > 0 && !string.IsNullOrEmpty(textePrecedent))
                CbTemps.Text = textePrecedent;
            else if (CbTemps.Items.Count > 0)
                CbTemps.Text = (CbTemps.Items[0] as ComboBoxItem)?.Content?.ToString() ?? "1";
            else
                CbTemps.Text = "";
        }

        private void OnValider(object sender, RoutedEventArgs e)
        {
            Fonction = (CbFonction.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Freq";
            NombrePlages = RbUnePlage.IsChecked == true ? 1 : 3;

            // Module sans temps de mesure : on saute toute la validation du temps.
            if (_module != null && !_module.UtiliseTempsDeMesure)
            {
                TempsDeMesure = 0;
                DialogResult = true;
                return;
            }

            string texte = (CbTemps.Text ?? "").Trim().Replace(',', '.');
            if (!double.TryParse(texte, NumberStyles.Float, CultureInfo.InvariantCulture, out double t) || t <= 0)
            {
                System.Windows.MessageBox.Show("Temps de mesure invalide.",
                    "Saisie incorrecte",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            // Refus si l'admin tape manuellement un temps déjà existant pour cette fonction.
            if (_module != null && _module.Lignes.Any(l =>
                string.Equals(l.Fonction, Fonction, System.StringComparison.OrdinalIgnoreCase) &&
                System.Math.Abs(l.TempsDeMesure - t) < 1e-9))
            {
                System.Windows.MessageBox.Show(
                    $"Le temps {t} s existe déjà pour la fonction « {Fonction} » dans ce module.\n" +
                    "Modifie les lignes existantes plutôt que d'en ajouter un doublon.",
                    "Doublon",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return;
            }

            TempsDeMesure = t;
            DialogResult = true;
        }

        private void OnAnnuler(object sender, RoutedEventArgs e) => DialogResult = false;
    }
}
