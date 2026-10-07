using System.Windows;
using System.Windows.Controls;
using Metrologo.Themes;

namespace Metrologo.Views
{
    public partial class SelectionUtilisateurView : UserControl
    {
        /// <summary>Marge de la ContentControl de MainWindow autour de la vue (16,16,16,8).</summary>
        private const double MargeHote = 16;

        public SelectionUtilisateurView()
        {
            InitializeComponent();
            AppliquerThemeSaisonnier();
        }

        /// <summary>
        /// Thème saisonnier en fond de l'écran de connexion, derrière la carte : bande en
        /// haut (guirlande, télécabines...), grand paysage en bas, neige ou pétales sur tout
        /// l'écran. Occupe le vide autour de la carte sans rien masquer. Hors saison : rien.
        /// </summary>
        private void AppliquerThemeSaisonnier()
        {
            var saison = DecorSaisonnier.Courante;
            if (saison == Saison.Aucune) return;

            // Insérés en tête : ils passent derrière la carte (premier enfant = tout au fond).
            var bande = DecorSaisonnier.Bande(saison);
            if (bande != null)
            {
                bande.VerticalAlignment = VerticalAlignment.Top;
                bande.Height = 52;
                bande.Margin = new Thickness(-MargeHote, -MargeHote, -MargeHote, 0);
                Racine.Children.Insert(0, bande);
            }

            var neige = DecorSaisonnier.Neige(saison);
            if (neige != null)
            {
                neige.Margin = new Thickness(-MargeHote, -MargeHote, -MargeHote, -8);
                Racine.Children.Insert(0, neige);
            }

            // Paysage en bas de l'écran : ~32 % de la hauteur, entre 110 et 340 px.
            var paysage = DecorSaisonnier.Paysage(saison, part: 0.32, mini: 110, maxi: 340);
            if (paysage != null)
            {
                paysage.Margin = new Thickness(-MargeHote, 0, -MargeHote, -8);
                Racine.Children.Insert(0, paysage);
            }
        }
    }
}
