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
        // Les temps de mesure standards qu'on propose à l'admin.
        private static readonly double[] _tempsStandards =
        {
            0.01, 0.02, 0.05, 0.1, 0.2, 0.5, 1d, 2d, 5d, 10d, 20d, 50d, 100d
        };

        private readonly ModuleIncertitude? _module;

        public string Fonction { get; private set; } = "Freq";
        public double TempsDeMesure { get; private set; } = 1.0;

        /// <summary>1 si on veut une seule plage, 3 pour le trio basse/moyenne/haute (par défaut).</summary>
        public int NombrePlages { get; private set; } = 3;

        public AjoutTrioDialog() : this(null) { }

        /// <summary>
        /// <paramref name="module"/> = le module qu'on est en train d'éditer. Quand il est
        /// fourni, on filtre la ComboBox des temps pour retirer ceux déjà saisis sur la
        /// fonction sélectionnée, histoire d'éviter les doublons.
        /// </summary>
        public AjoutTrioDialog(ModuleIncertitude? module)
        {
            InitializeComponent();
            _module = module;

            // Pour un module sans temps de mesure, on cache tout le bloc et OnValider mettra
            // TempsDeMesure à 0 : choisir un temps n'aurait aucun sens (tachy/strobo). Au
            // passage on coche « 1 plage », qui est le cas habituel en tachy.
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

            // On met à jour le libellé du bouton au fil du choix de l'utilisateur.
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

            // Les temps déjà présents dans le tableau pour cette fonction, sur le module courant.
            HashSet<double> dejaPresents = _module == null
                ? new HashSet<double>()
                : new HashSet<double>(_module.Lignes
                    .Where(l => string.Equals(l.Fonction, fnSel, System.StringComparison.OrdinalIgnoreCase))
                    .Select(l => l.TempsDeMesure));

            // On mémorise le texte courant pour ne pas écraser une saisie en cours.
            string textePrecedent = CbTemps.Text;

            CbTemps.Items.Clear();
            foreach (double t in _tempsStandards)
            {
                if (dejaPresents.Contains(t)) continue; // on saute, il est déjà là
                CbTemps.Items.Add(new ComboBoxItem
                {
                    Content = t.ToString(CultureInfo.InvariantCulture)
                });
            }

            // Si la liste n'est pas vide, on sélectionne le premier temps restant.
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

            // Module sans temps de mesure : inutile de valider quoi que ce soit côté temps.
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

            // On refuse si l'admin retape à la main un temps qui existe déjà pour cette fonction.
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
