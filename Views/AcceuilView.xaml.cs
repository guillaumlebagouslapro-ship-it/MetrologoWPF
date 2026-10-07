using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using Metrologo.Themes;
using Metrologo.ViewModels;

namespace Metrologo.Views
{
    public partial class AcceuilView : UserControl
    {
        /// <summary>Marge de la ContentControl de MainWindow autour de la vue (16,16,16,8).</summary>
        private const double MargeHote = 16;

        /// <summary>Place libérée en haut pour la bande saisonnière (guirlande, télécabines...).</summary>
        private const double PlaceBande = 32;

        public AcceuilView()
        {
            InitializeComponent();

            string? version = UpdateService.VersionInstallee();
            if (version != null)
                MisesAJourTexte.Text = "Version " + HistoriqueMaj.Affichage(version);

            AppliquerThemeSaisonnier();
        }

        /// <summary>
        /// Décorations du thème saisonnier (Noël, ski, été), choisi d'après la date du jour
        /// (voir <see cref="DecorSaisonnier.Determiner"/>). Hors saison : rien n'est ajouté.
        /// Les calques débordent de la vue (marges négatives) pour couvrir toute la zone
        /// sous la barre de navigation.
        /// </summary>
        private void AppliquerThemeSaisonnier()
        {
            var saison = DecorSaisonnier.Courante;
            if (saison == Saison.Aucune) return;

            PastilleSaison.Content = DecorSaisonnier.Pastille(saison);
            FondBandeauSaison.Content = DecorSaisonnier.FondBandeau(saison, sombre: true, rayon: 10);

            int lignes = Racine.RowDefinitions.Count;

            var neige = DecorSaisonnier.Neige(saison);
            if (neige != null)
            {
                Grid.SetRowSpan(neige, lignes);
                Panel.SetZIndex(neige, 50);
                neige.Margin = new Thickness(-MargeHote, -MargeHote - PlaceBande, -MargeHote, -8);
                Racine.Children.Add(neige);
            }

            var bande = DecorSaisonnier.Bande(saison);
            if (bande != null)
            {
                Racine.Margin = new Thickness(0, PlaceBande, 0, 0);
                Grid.SetRowSpan(bande, lignes);
                Panel.SetZIndex(bande, 51);
                bande.VerticalAlignment = VerticalAlignment.Top;
                bande.Height = 52;
                bande.Margin = new Thickness(-MargeHote, -MargeHote - PlaceBande, -MargeHote, 0);
                Racine.Children.Add(bande);
            }

            // L'accessoire vit dans le gabarit du bouton « Lancer la mesure ».
            BoutonLancer.Loaded += (_, _) =>
            {
                if (BoutonLancer.Template?.FindName("AccessoireSaison", BoutonLancer) is ContentControl cc
                    && cc.Content == null)
                    cc.Content = DecorSaisonnier.Accessoire(saison);
            };

            // Animations figées pendant une mesure : rien ne bouge à l'écran quand on lit les valeurs.
            DataContextChanged += (_, e) =>
            {
                if (e.OldValue is INotifyPropertyChanged ancien) ancien.PropertyChanged -= SuivreMesure;
                if (e.NewValue is INotifyPropertyChanged nouveau) nouveau.PropertyChanged += SuivreMesure;
            };
        }

        private void SuivreMesure(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != nameof(AccueilViewModel.MesureEnCours) || sender is not AccueilViewModel vm) return;
            Dispatcher.BeginInvoke(() =>
            {
                if (vm.MesureEnCours) DecorSaisonnier.Pause();
                else DecorSaisonnier.Reprendre();
            });
        }

        /// <summary>Accueil > Mises à jour : historique + retour à la version précédente.</summary>
        private void MisesAJour_Click(object sender, RoutedEventArgs e)
        {
            new MisesAJourWindow { Owner = Window.GetWindow(this) }.ShowDialog();
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
