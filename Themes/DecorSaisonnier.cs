using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using Forme = System.Windows.Shapes;

namespace Metrologo.Themes;

/// <summary>Thème saisonnier affiché sur l'écran principal.</summary>
public enum Saison { Aucune, Noel, Ski, Paques, Ete }

/// <summary>
/// Décorations saisonnières discrètes (guirlande, neige, télécabines, soleil...).
/// <para/>
/// La saison est choisie au démarrage d'après la date du jour, chaque année :
/// Noël en décembre, ski en janvier-février, Pâques du lundi de Pâques pendant deux
/// semaines (date calculée), été en juillet-août, rien le reste du temps. Pour tester un
/// thème avant sa date : variable d'environnement
/// <c>ASERTI_THEME</c> = noel | ski | paques | ete | aucun.
/// <para/>
/// Tout est purement visuel : non cliquable (IsHitTestVisible = false), animations
/// limitées à 30 images/s et suspendues d'un coup par <see cref="Pause"/> (pendant une
/// mesure). Fichier identique dans Asertools (seul le namespace change).
/// </summary>
public static class DecorSaisonnier
{
    public const string VariableForcage = "ASERTI_THEME";

    /// <summary>Saison de cette exécution (date du démarrage, ou forçage).</summary>
    public static Saison Courante { get; } = Determiner(DateTime.Today);

    /// <summary>Calendrier des thèmes. Modifier ici pour changer les périodes.</summary>
    public static Saison Determiner(DateTime jour)
    {
        string? force = Environment.GetEnvironmentVariable(VariableForcage);
        if (!string.IsNullOrWhiteSpace(force))
        {
            return force.Trim().ToLowerInvariant() switch
            {
                "noel" or "noël" => Saison.Noel,
                "ski" => Saison.Ski,
                "paques" or "pâques" => Saison.Paques,
                "ete" or "été" => Saison.Ete,
                _ => Saison.Aucune,
            };
        }

        // Pâques : du lundi de Pâques inclus, pendant 14 jours.
        DateTime lundiPaques = DimanchePaques(jour.Year).AddDays(1);
        if (jour.Date >= lundiPaques && jour.Date < lundiPaques.AddDays(14))
            return Saison.Paques;

        return jour.Month switch
        {
            12 => Saison.Noel,      // 1er -> 31 décembre
            1 or 2 => Saison.Ski,   // 1er janvier -> fin février
            7 or 8 => Saison.Ete,   // 1er juillet -> 31 août
            _ => Saison.Aucune,
        };
    }

    /// <summary>Dimanche de Pâques (calendrier grégorien, algorithme de Meeus/Jones/Butcher).</summary>
    public static DateTime DimanchePaques(int annee)
    {
        int a = annee % 19, b = annee / 100, c = annee % 100;
        int d = b / 4, e = b % 4, f = (b + 8) / 25, g = (b - f + 1) / 3;
        int h = (19 * a + b - d - g + 15) % 30;
        int i = c / 4, k = c % 4;
        int l = (32 + 2 * e + 2 * i - h - k) % 7;
        int m = (a + 11 * h + 22 * l) / 451;
        int mois = (h + l - 7 * m + 114) / 31;
        int jour = (h + l - 7 * m + 114) % 31 + 1;
        return new DateTime(annee, mois, jour);
    }

    // ------------------------------------------------------------------ pause globale

    private static readonly List<AnimationClock> Horloges = new();
    private static bool _enPause;

    /// <summary>Fige toutes les animations (ex. pendant une mesure).</summary>
    public static void Pause()
    {
        _enPause = true;
        foreach (var h in Horloges) h.Controller?.Pause();
    }

    public static void Reprendre()
    {
        _enPause = false;
        foreach (var h in Horloges) h.Controller?.Resume();
    }

    // ------------------------------------------------------------------ éléments

    /// <summary>Neige qui tombe sur toute la zone (Noël, ski), pétales pastel (Pâques). Null sinon.</summary>
    public static FrameworkElement? Neige(Saison s)
    {
        if (s == Saison.Paques) return new Scene(Petales);
        if (s != Saison.Noel && s != Saison.Ski) return null;
        bool ski = s == Saison.Ski;
        return new Scene((sc, l, h) =>
        {
            var r = new Alea(ski ? 11 : 7);
            int n = (int)Math.Round((ski ? 18 : 34) * Math.Max(0.5, l / 1180.0));
            for (int i = 0; i < n; i++)
            {
                double taille = ski ? 2.5 + r.Suivant() * 4 : 3 + r.Suivant() * 5;
                double duree = ski ? 14 + r.Suivant() * 12 : 10 + r.Suivant() * 10;
                var flocon = new Forme.Ellipse
                {
                    Width = taille,
                    Height = taille,
                    Fill = Brushes.White,
                    Stroke = Pinceau("#B5C4D8"),
                    StrokeThickness = 0.6,
                    Opacity = 0.55 + r.Suivant() * 0.4,
                };
                Canvas.SetLeft(flocon, r.Suivant() * l);
                var tt = new TranslateTransform();
                flocon.RenderTransform = tt;
                sc.Children.Add(flocon);

                var debut = TimeSpan.FromSeconds(-r.Suivant() * duree);
                double derive = (r.Suivant() < 0.5 ? -1 : 1) * (ski ? 30 : 46);
                sc.Animer(tt, TranslateTransform.YProperty, Boucle(-12, h + 12, duree, debut));
                sc.Animer(tt, TranslateTransform.XProperty, Boucle(0, derive, duree, debut));
            }
        });
    }

    /// <summary>Bande décorative à poser sous une barre de navigation / un en-tête :
    /// guirlande (Noël), câble + télécabines (ski), fanions pastel (Pâques), mouettes (été).
    /// Hauteur conseillée : 52.</summary>
    public static FrameworkElement? Bande(Saison s) => s switch
    {
        Saison.Noel => new Scene(Guirlande),
        Saison.Ski => new Scene(Telecabines),
        Saison.Paques => new Scene(Fanions),
        Saison.Ete => new Scene(Mouettes),
        _ => null,
    };

    /// <summary>Décor de fond d'un bandeau ou en-tête (congère, montagnes, prairie + œufs, soleil + vagues),
    /// découpé aux coins arrondis <paramref name="rayon"/>. <paramref name="sombre"/> = fond
    /// marine (Metrologo) ou blanc (Asertools).</summary>
    public static FrameworkElement? FondBandeau(Saison s, bool sombre, double rayon)
    {
        if (s == Saison.Aucune || (s == Saison.Noel && !sombre)) return null;
        return new Scene((sc, l, h) =>
        {
            sc.Clip = new RectangleGeometry(new Rect(0, 0, l, h), rayon, rayon);
            switch (s)
            {
                case Saison.Noel: Congere(sc, l); break;
                case Saison.Ski: Montagnes(sc, l, h, sombre); break;
                case Saison.Paques: Prairie(sc, l, h, sombre); break;
                case Saison.Ete: SoleilEtVagues(sc, l, h, sombre); break;
            }
        });
    }

    /// <summary>Paysage de saison à poser en bas d'une zone (logs, écran de connexion) :
    /// sapins enneigés et bonhomme de neige, piste avec skieur, prairie fleurie avec lapin,
    /// plage avec voilier. Proportions calées sur 130 px de haut, s'adapte à toute taille.
    /// <paramref name="rayon"/> = coins arrondis du cadre qui l'accueille.
    /// <para/>
    /// À poser en pleine zone (Stretch) : le paysage occupe le bas, sur une hauteur
    /// proportionnelle à la zone (<paramref name="part"/>), bornée entre
    /// <paramref name="mini"/> et <paramref name="maxi"/> et limitée par la largeur. Les
    /// éléments sont mis à l'échelle uniformément (jamais déformés) : lisible sur un grand
    /// écran, sans envahir un petit.</summary>
    public static FrameworkElement? Paysage(Saison s, double rayon = 0, double part = 0.45,
                                            double mini = 60, double maxi = 220)
    {
        if (s == Saison.Aucune) return null;
        var scene = new Scene((sc, l, h) =>
        {
            if (rayon > 0) sc.Clip = new RectangleGeometry(new Rect(0, 0, l, h), rayon, rayon);
            double k = h / 130.0;
            switch (s)
            {
                case Saison.Noel: PaysageNoel(sc, l, h, k); break;
                case Saison.Ski: PaysageSki(sc, l, h, k); break;
                case Saison.Paques: PaysagePaques(sc, l, h, k); break;
                case Saison.Ete: PaysageEte(sc, l, h, k); break;
            }
        }) { VerticalAlignment = VerticalAlignment.Bottom };

        var zone = new Grid { IsHitTestVisible = false };
        zone.Children.Add(scene);
        zone.SizeChanged += (_, _) =>
        {
            double voulu = Math.Min(zone.ActualHeight * part, zone.ActualWidth * 0.28);
            scene.Height = Math.Max(Math.Min(mini, zone.ActualHeight), Math.Min(voulu, maxi));
        };
        return zone;
    }

    /// <summary>Petit accessoire posé sur une icône carrée de 44 px (bonnet de Noël,
    /// bonnet de ski, lunettes de soleil). Déjà positionné (marges négatives).</summary>
    public static FrameworkElement? Accessoire(Saison s)
    {
        var c = new Canvas { IsHitTestVisible = false, VerticalAlignment = VerticalAlignment.Top };
        switch (s)
        {
            case Saison.Noel:
                c.Width = 36; c.Height = 30;
                c.Children.Add(Chemin("M3 24 Q 12 2 27 6 Q 31 8 31 14 L 26 12 Q 20 12 15 24 Z", "#D93A3A"));
                c.Children.Add(Rect(1, 21, 17, 6, 3, "#FFFFFF"));
                c.Children.Add(Rond(31, 15, 3.6, "#FFFFFF"));
                c.HorizontalAlignment = HorizontalAlignment.Right;
                c.Margin = new Thickness(0, -17, -16, 0);
                c.RenderTransform = new RotateTransform(18, 18, 15);
                return c;
            case Saison.Ski:
                c.Width = 36; c.Height = 32;
                c.Children.Add(Chemin("M5 24 Q 5 8 18 8 Q 31 8 31 24 Z", "#3A7BD5"));
                c.Children.Add(Rect(3, 21, 30, 7, 3, "#FFFFFF"));
                var tirets = Chemin("M7 24.5 L29 24.5", null);
                tirets.Stroke = Pinceau("#3A7BD5");
                tirets.StrokeThickness = 1.6;
                tirets.StrokeDashArray = new DoubleCollection { 1.25, 1.25 };
                c.Children.Add(tirets);
                var pompon = Rond(18, 6, 4.5, "#FFFFFF");
                pompon.Stroke = Pinceau("#C9D6E8");
                pompon.StrokeThickness = 1;
                c.Children.Add(pompon);
                c.HorizontalAlignment = HorizontalAlignment.Right;
                c.Margin = new Thickness(0, -19, -14, 0);
                c.RenderTransform = new RotateTransform(14, 18, 16);
                return c;
            case Saison.Paques:
                // Serre-tête à oreilles de lapin, centré sur l'icône.
                c.Width = 36; c.Height = 30;
                var bandeau = Chemin("M4 29 Q 18 19 32 29", null);
                bandeau.Stroke = Pinceau("#C3B1E1");
                bandeau.StrokeThickness = 3;
                bandeau.StrokeStartLineCap = PenLineCap.Round;
                bandeau.StrokeEndLineCap = PenLineCap.Round;
                c.Children.Add(Oreille(7, 1, -14));
                c.Children.Add(Oreille(19, 1, 14));
                c.Children.Add(bandeau);
                c.HorizontalAlignment = HorizontalAlignment.Center;
                c.Margin = new Thickness(0, -25, 0, 0);
                return c;
            case Saison.Ete:
                c.Width = 40; c.Height = 16;
                c.Children.Add(Rect(2, 3, 15, 11, 5, "#1F2937"));
                c.Children.Add(Rect(23, 3, 15, 11, 5, "#1F2937"));
                var pont = Chemin("M17 7 Q 20 4 23 7", null);
                pont.Stroke = Pinceau("#1F2937");
                pont.StrokeThickness = 2;
                c.Children.Add(pont);
                var reflet = Chemin("M5 6.5 L10 6.5 M26 6.5 L31 6.5", null);
                reflet.Stroke = Brushes.White;
                reflet.StrokeThickness = 1.4;
                reflet.Opacity = 0.7;
                c.Children.Add(reflet);
                c.HorizontalAlignment = HorizontalAlignment.Left;
                c.Margin = new Thickness(2, -9, 0, 0);
                c.RenderTransform = new RotateTransform(-6, 20, 8);
                return c;
            default:
                return null;
        }
    }

    /// <summary>Pastille de texte (« Joyeuses fêtes », « Bonne année 2027 », « Bel été ! »).</summary>
    public static FrameworkElement? Pastille(Saison s)
    {
        (string texte, string fond, string encre) = s switch
        {
            Saison.Noel => ("Joyeuses fêtes", "#FBE7E7", "#A1231B"),
            Saison.Ski => ($"Bonne année {DateTime.Today.Year}", "#E2EFFA", "#12568A"),
            Saison.Paques => ("Joyeuses Pâques", "#EFE7FA", "#5B2B91"),
            Saison.Ete => ("Bel été !", "#FFEFD6", "#8A4B00"),
            _ => ("", "", ""),
        };
        if (texte.Length == 0) return null;
        return new Border
        {
            Background = Pinceau(fond),
            CornerRadius = new CornerRadius(10),
            Padding = new Thickness(10, 2, 10, 3),
            Margin = new Thickness(10, 0, 0, 0),
            VerticalAlignment = VerticalAlignment.Center,
            IsHitTestVisible = false,
            Child = new TextBlock
            {
                Text = texte,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = Pinceau(encre),
            },
        };
    }

    // ------------------------------------------------------------------ scènes

    private static void Guirlande(Scene sc, double l, double h)
    {
        string[] couleurs = { "#D93A3A", "#F2B33D", "#2F9E63", "#3A7BD5" };
        int festons = Math.Max(3, (int)Math.Round(l / 195));
        const double creux = 11;
        const int parFeston = 4;
        double pas = l / festons;

        var d = new StringBuilder("M 0 2");
        var ampoules = new List<(double x, double y)>();
        for (int i = 0; i < festons; i++)
        {
            double x0 = i * pas, x1 = x0 + pas, xm = x0 + pas / 2;
            d.Append(F($" Q {xm} {2 + 2 * creux} {x1} 2"));
            for (int k = 1; k <= parFeston; k++)
            {
                double t = k / (double)(parFeston + 1);
                double x = (1 - t) * (1 - t) * x0 + 2 * (1 - t) * t * xm + t * t * x1;
                double y = 2 + 4 * (1 - t) * t * creux;
                ampoules.Add((x, y));
            }
        }
        var fil = Chemin(d.ToString(), null);
        fil.Stroke = Pinceau("#2F6B45");
        fil.StrokeThickness = 2;
        sc.Children.Add(fil);

        for (int n = 0; n < ampoules.Count; n++)
        {
            var (x, y) = ampoules[n];
            var couleur = (Color)ColorConverter.ConvertFromString(couleurs[n % couleurs.Length]);
            var groupe = new Canvas();
            var halo = new Forme.Ellipse
            {
                Width = 20,
                Height = 20,
                Fill = new RadialGradientBrush(Color.FromArgb(150, couleur.R, couleur.G, couleur.B),
                                               Color.FromArgb(0, couleur.R, couleur.G, couleur.B)),
            };
            Canvas.SetLeft(halo, -6); Canvas.SetTop(halo, -2);
            groupe.Children.Add(halo);
            groupe.Children.Add(Rect(2, -2, 4, 4, 1, "#6B7488"));
            var ampoule = new Forme.Ellipse { Width = 8, Height = 11, Fill = new SolidColorBrush(couleur) };
            Canvas.SetTop(ampoule, 1);
            groupe.Children.Add(ampoule);
            Canvas.SetLeft(groupe, x - 4);
            Canvas.SetTop(groupe, y + 1);
            sc.Children.Add(groupe);

            var clignote = new DoubleAnimation(1, 0.35, new Duration(TimeSpan.FromSeconds(2.2)))
            {
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                BeginTime = TimeSpan.FromSeconds(-((n * 0.37) % 2.4)),
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
            };
            sc.Animer(groupe, UIElement.OpacityProperty, clignote);
        }
    }

    private static void Telecabines(Scene sc, double l, double h)
    {
        var cable = new Forme.Line { X1 = 0, Y1 = 3, X2 = l, Y2 = 3, Stroke = Pinceau("#3B4A63"), StrokeThickness = 1.6 };
        sc.Children.Add(cable);

        string[] couleurs = { "#D93A3A", "#E89B2F" };
        int n = l > 800 ? 2 : 1;
        double duree = 34 * l / 1180.0 + 6;
        for (int i = 0; i < n; i++)
        {
            var cabine = new Canvas { Width = 44, Height = 44 };
            cabine.Children.Add(Rond(22, 3, 3, "#3B4A63"));
            var tige = new Forme.Line { X1 = 22, Y1 = 3, X2 = 22, Y2 = 11, Stroke = Pinceau("#3B4A63"), StrokeThickness = 2 };
            cabine.Children.Add(tige);
            cabine.Children.Add(Rect(6, 11, 32, 26, 6, couleurs[i % couleurs.Length]));
            cabine.Children.Add(Rect(10, 15, 10, 9, 2, "#E8F2FB"));
            cabine.Children.Add(Rect(24, 15, 10, 9, 2, "#E8F2FB"));
            var bas = Rect(6, 31, 32, 3, 0, "#1E2D4F");
            bas.Opacity = 0.25;
            cabine.Children.Add(bas);

            var balance = new RotateTransform(0, 22, 3);
            var deplacement = new TranslateTransform();
            var groupe = new TransformGroup();
            groupe.Children.Add(balance);
            groupe.Children.Add(deplacement);
            cabine.RenderTransform = groupe;
            sc.Children.Add(cabine);

            sc.Animer(deplacement, TranslateTransform.XProperty,
                Boucle(-60, l + 20, duree, TimeSpan.FromSeconds(-i * duree / n)));
            sc.Animer(balance, RotateTransform.AngleProperty, new DoubleAnimation(-3, 3, new Duration(TimeSpan.FromSeconds(2.6)))
            {
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
            });
        }
    }

    private static readonly string[] Pastels = { "#F4A7B9", "#F7D774", "#9ED9C3", "#C3B1E1", "#A7CDF2" };

    private static void Petales(Scene sc, double l, double h)
    {
        var r = new Alea(5);
        int n = (int)Math.Round(14 * Math.Max(0.5, l / 1180.0));
        for (int i = 0; i < n; i++)
        {
            double duree = 16 + r.Suivant() * 10;
            var petale = new Forme.Ellipse
            {
                Width = 7,
                Height = 4.5,
                Fill = Pinceau(i % 3 == 0 ? "#FFFFFF" : "#F4A7B9"),
                Stroke = Pinceau("#E9B7C6"),
                StrokeThickness = 0.5,
                Opacity = 0.75,
            };
            Canvas.SetLeft(petale, r.Suivant() * l);
            var rotation = new RotateTransform(0, 3.5, 2.25);
            var chute = new TranslateTransform();
            var groupe = new TransformGroup();
            groupe.Children.Add(rotation);
            groupe.Children.Add(chute);
            petale.RenderTransform = groupe;
            sc.Children.Add(petale);

            var debut = TimeSpan.FromSeconds(-r.Suivant() * duree);
            sc.Animer(chute, TranslateTransform.YProperty, Boucle(-10, h + 10, duree, debut));
            sc.Animer(chute, TranslateTransform.XProperty, Boucle(0, (r.Suivant() < 0.5 ? -1 : 1) * 60, duree, debut));
            sc.Animer(rotation, RotateTransform.AngleProperty, Boucle(0, r.Suivant() < 0.5 ? 360 : -360, 4 + r.Suivant() * 4, debut));
        }
    }

    private static void Fanions(Scene sc, double l, double h)
    {
        int festons = Math.Max(2, (int)Math.Round(l / 300));
        const double creux = 9;
        const int parFeston = 6;
        double pas = l / festons;

        var d = new StringBuilder("M 0 2");
        int n = 0;
        var fanions = new List<(double x, double y)>();
        for (int i = 0; i < festons; i++)
        {
            double x0 = i * pas, x1 = x0 + pas, xm = x0 + pas / 2;
            d.Append(F($" Q {xm} {2 + 2 * creux} {x1} 2"));
            for (int k = 1; k <= parFeston; k++)
            {
                double t = k / (double)(parFeston + 1);
                fanions.Add(((1 - t) * (1 - t) * x0 + 2 * (1 - t) * t * xm + t * t * x1,
                             2 + 4 * (1 - t) * t * creux));
            }
        }
        var fil = Chemin(d.ToString(), null);
        fil.Stroke = Pinceau("#B9A5D6");
        fil.StrokeThickness = 1.5;
        sc.Children.Add(fil);

        foreach (var (x, y) in fanions)
        {
            var fanion = Chemin("M0 0 L14 0 L7 15 Z", Pastels[n % Pastels.Length]);
            Canvas.SetLeft(fanion, x - 7);
            Canvas.SetTop(fanion, y);
            var balance = new RotateTransform(0, 7, 0);
            fanion.RenderTransform = balance;
            sc.Children.Add(fanion);
            sc.Animer(balance, RotateTransform.AngleProperty, new DoubleAnimation(-5, 5, new Duration(TimeSpan.FromSeconds(3)))
            {
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                BeginTime = TimeSpan.FromSeconds(-((n * 0.53) % 3)),
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
            });
            n++;
        }
    }

    private static void Prairie(Scene sc, double l, double h, bool sombre)
    {
        double xLapin = Math.Max(l * 0.38, l - (sombre ? 470 : 340));

        // Oreilles de lapin qui sortent de l'herbe de temps en temps (dessinées avant l'herbe).
        var lapin = new Canvas();
        lapin.Children.Add(Oreille(0, 0, -10));
        lapin.Children.Add(Oreille(13, 0, 10));
        Canvas.SetLeft(lapin, xLapin);
        Canvas.SetTop(lapin, h - 32);
        var montee = new TranslateTransform(0, 28);
        lapin.RenderTransform = montee;
        sc.Children.Add(lapin);
        var coucou = new DoubleAnimationUsingKeyFrames { Duration = TimeSpan.FromSeconds(11), RepeatBehavior = RepeatBehavior.Forever };
        coucou.KeyFrames.Add(new LinearDoubleKeyFrame(28, KeyTime.FromTimeSpan(TimeSpan.Zero)));
        coucou.KeyFrames.Add(new LinearDoubleKeyFrame(28, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(5))));
        coucou.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(5.7)), new SineEase()));
        coucou.KeyFrames.Add(new LinearDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(8))));
        coucou.KeyFrames.Add(new EasingDoubleKeyFrame(28, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(8.7)), new SineEase()));
        sc.Animer(montee, TranslateTransform.YProperty, coucou);

        // Herbe : brins en dents de scie sur toute la largeur.
        var herbe = new StringBuilder(F($"M 0 {h}"));
        var r = new Alea(13);
        for (double x = 0; x < l; x += 9)
        {
            double haut = 7 + r.Suivant() * 7;
            herbe.Append(F($" L {x + 4.5} {h - haut} L {x + 9} {h - 2}"));
        }
        herbe.Append(F($" L {l} {h} Z"));
        sc.Children.Add(Chemin(herbe.ToString(), sombre ? "#4E9A5B" : "#9ED39A"));

        // Œufs décorés posés dans l'herbe.
        double[] decalages = { 60, 98, 124 };
        for (int i = 0; i < decalages.Length; i++)
        {
            var oeuf = new Canvas { Width = 15, Height = 19 };
            oeuf.Children.Add(new Forme.Ellipse { Width = 15, Height = 19, Fill = Pinceau(Pastels[(i * 2) % Pastels.Length]) });
            var motif = Chemin("M1.5 9 L4 7 L7.5 9.5 L11 7 L13.5 9", null);
            motif.Stroke = Brushes.White;
            motif.StrokeThickness = 1.4;
            oeuf.Children.Add(motif);
            Canvas.SetLeft(oeuf, xLapin + decalages[i]);
            Canvas.SetTop(oeuf, h - 21 + (i == 1 ? 2 : 0));
            oeuf.RenderTransform = new RotateTransform(i == 1 ? 12 : -8, 7.5, 9.5);
            sc.Children.Add(oeuf);
        }
    }

    // ------------------------------------------------------------------ paysages

    private static void PaysageNoel(Scene sc, double l, double h, double k)
    {
        // Étoiles qui scintillent dans le ciel.
        var r = new Alea(29);
        for (int i = 0; i < 9; i++)
        {
            double sx = (0.04 + r.Suivant() * 0.92) * l, sy = (8 + r.Suivant() * 34) * k, sr = (2.2 + r.Suivant() * 1.6) * k;
            var etoile = Chemin(F($"M {sx} {sy - sr * 2} L {sx + sr * 0.5} {sy - sr * 0.5} L {sx + sr * 2} {sy} L {sx + sr * 0.5} {sy + sr * 0.5} L {sx} {sy + sr * 2} L {sx - sr * 0.5} {sy + sr * 0.5} L {sx - sr * 2} {sy} L {sx - sr * 0.5} {sy - sr * 0.5} Z"), "#F2C14E");
            sc.Children.Add(etoile);
            sc.Animer(etoile, UIElement.OpacityProperty, new DoubleAnimation(0.9, 0.2, new Duration(TimeSpan.FromSeconds(1.6 + r.Suivant() * 1.4)))
            {
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                BeginTime = TimeSpan.FromSeconds(-r.Suivant() * 3),
            });
        }

        Collines(sc, l, h, 62 * k, 10 * k, 3, "#DCE6F2", 1, null);

        // Chalet avec fenêtre éclairée.
        double cx0 = 0.52 * l, cy0 = h - 44 * k;
        sc.Children.Add(Rect(cx0 - 22 * k, cy0 - 26 * k, 44 * k, 26 * k, 1, "#8B5E3C"));
        sc.Children.Add(Rect(cx0 + 10 * k, cy0 - 48 * k, 7 * k, 14 * k, 0, "#6B4A30"));
        sc.Children.Add(Chemin(F($"M {cx0 - 28 * k} {cy0 - 24 * k} L {cx0} {cy0 - 46 * k} L {cx0 + 28 * k} {cy0 - 24 * k} Z"), "#6B4A30"));
        sc.Children.Add(Chemin(F($"M {cx0 - 29 * k} {cy0 - 23 * k} L {cx0} {cy0 - 47 * k} L {cx0 + 29 * k} {cy0 - 23 * k} L {cx0 + 22 * k} {cy0 - 23 * k} L {cx0} {cy0 - 40 * k} L {cx0 - 22 * k} {cy0 - 23 * k} Z"), "#FFFFFF"));
        sc.Children.Add(Rect(cx0 - 4 * k, cy0 - 14 * k, 8 * k, 14 * k, 1, "#5A3D26"));
        var fenetre = Rect(cx0 - 17 * k, cy0 - 19 * k, 9 * k, 8 * k, 1, "#F7D774");
        sc.Children.Add(fenetre);
        sc.Children.Add(Rect(cx0 + 8 * k, cy0 - 19 * k, 9 * k, 8 * k, 1, "#F7D774"));
        sc.Animer(fenetre, UIElement.OpacityProperty, new DoubleAnimation(1, 0.65, new Duration(TimeSpan.FromSeconds(2.5)))
        {
            AutoReverse = true,
            RepeatBehavior = RepeatBehavior.Forever,
        });

        double[] xs = { 0.05, 0.11, 0.17, 0.42, 0.61, 0.70, 0.77, 0.86, 0.94 };
        double[] ts = { 70, 52, 60, 50, 58, 64, 82, 56, 68 };
        const int grandSapin = 6;
        for (int i = 0; i < xs.Length; i++)
            sc.Children.Add(Sapin(xs[i] * l, h - 38 * k, ts[i] * k, i % 2 == 0 ? "#2F6B45" : "#3E7D55", true));

        // Guirlande lumineuse sur le grand sapin.
        string[] couleurs = { "#D93A3A", "#F2B33D", "#3A7BD5", "#2F9E63" };
        (double dx, double dy)[] lumieres = { (-0.2, 0.22), (0.16, 0.3), (-0.08, 0.42), (0.15, 0.5), (-0.12, 0.62), (0.07, 0.72), (-0.02, 0.86) };
        double t = ts[grandSapin] * k, cx = xs[grandSapin] * l, pied = h - 38 * k;
        for (int i = 0; i < lumieres.Length; i++)
        {
            var lampe = Rond(cx + lumieres[i].dx * t, pied - lumieres[i].dy * t, 2.4 * k, couleurs[i % couleurs.Length]);
            sc.Children.Add(lampe);
            sc.Animer(lampe, UIElement.OpacityProperty, new DoubleAnimation(1, 0.3, new Duration(TimeSpan.FromSeconds(1.8)))
            {
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                BeginTime = TimeSpan.FromSeconds(-i * 0.45),
            });
        }

        Collines(sc, l, h, 38 * k, 6 * k, 4, "#FFFFFF", 0, "#C9D6E5");

        // Bonhomme de neige.
        double bx = 0.33 * l, by = h - 24 * k;
        var corps = Rond(bx, by - 14 * k, 15 * k, "#FFFFFF");
        corps.Stroke = Pinceau("#C9D6E5"); corps.StrokeThickness = 1;
        var tete = Rond(bx, by - 37 * k, 10 * k, "#FFFFFF");
        tete.Stroke = Pinceau("#C9D6E5"); tete.StrokeThickness = 1;
        sc.Children.Add(corps);
        sc.Children.Add(tete);
        sc.Children.Add(Rect(bx - 10 * k, by - 29 * k, 20 * k, 4.5 * k, 2 * k, "#D93A3A"));
        sc.Children.Add(Rond(bx - 3.5 * k, by - 39 * k, 1.3 * k, "#1F2937"));
        sc.Children.Add(Rond(bx + 3.5 * k, by - 39 * k, 1.3 * k, "#1F2937"));
        sc.Children.Add(Chemin(F($"M {bx} {by - 36 * k} L {bx + 9 * k} {by - 35 * k} L {bx} {by - 34 * k} Z"), "#E8902F"));
        sc.Children.Add(Rond(bx, by - 17 * k, 1.4 * k, "#1F2937"));
        sc.Children.Add(Rond(bx, by - 9 * k, 1.4 * k, "#1F2937"));
        sc.Children.Add(Rect(bx - 7 * k, by - 57 * k, 14 * k, 11 * k, 1 * k, "#1F2937"));
        sc.Children.Add(Rect(bx - 10.5 * k, by - 47.5 * k, 21 * k, 2.6 * k, 1 * k, "#1F2937"));
    }

    private static void PaysageSki(Scene sc, double l, double h, double k)
    {
        double[] px = { 0.12, 0.42, 0.70 }, py = { 95, 120, 105 };
        sc.Children.Add(Chemin(F($"M 0 {h} L 0 {h - 60 * k} L {px[0] * l} {h - py[0] * k} L {0.25 * l} {h - 70 * k} L {px[1] * l} {h - py[1] * k} L {0.58 * l} {h - 78 * k} L {px[2] * l} {h - py[2] * k} L {0.86 * l} {h - 66 * k} L {l} {h - 90 * k} L {l} {h} Z"), "#D6E2EF"));
        for (int i = 0; i < px.Length; i++)
        {
            double x = px[i] * l, y = h - py[i] * k;
            sc.Children.Add(Chemin(F($"M {x - 14 * k} {y + 16 * k} L {x} {y} L {x + 14 * k} {y + 16 * k} L {x + 5 * k} {y + 12 * k} L {x} {y + 17 * k} L {x - 5 * k} {y + 12 * k} Z"), "#FFFFFF"));
        }

        // Piste : courbe y(t) avec x = l·t (point de contrôle centré).
        double ya = h - 30 * k, yc = h - 70 * k, yb = h - 24 * k;
        var piste = Chemin(F($"M 0 {h} L 0 {ya} Q {0.5 * l} {yc} {l} {yb} L {l} {h} Z"), "#F4F8FC");
        piste.Stroke = Pinceau("#C9D6E5");
        piste.StrokeThickness = 1;
        sc.Children.Add(piste);
        double[] xs = { 0.03, 0.08, 0.92, 0.97 };
        for (int i = 0; i < xs.Length; i++)
            sc.Children.Add(Sapin(xs[i] * l, h - 4 * k, (i % 2 == 0 ? 62 : 48) * k, "#3E6B5A", true));

        // Skieur qui descend la piste.
        var skieur = new Canvas();
        var ski = Chemin("M -11 1 L 13 -1", null);
        ski.Stroke = Pinceau("#1E2D4F"); ski.StrokeThickness = 2; ski.StrokeStartLineCap = PenLineCap.Round; ski.StrokeEndLineCap = PenLineCap.Round;
        var corps = Chemin("M 0 0 L 3 -9 L 0 -18", null);
        corps.Stroke = Pinceau("#3A7BD5"); corps.StrokeThickness = 4; corps.StrokeLineJoin = PenLineJoin.Round; corps.StrokeStartLineCap = PenLineCap.Round;
        var baton = Chemin("M 1 -15 L 9 -9 L 14 0", null);
        baton.Stroke = Pinceau("#1E2D4F"); baton.StrokeThickness = 1.2;
        skieur.Children.Add(ski);
        skieur.Children.Add(corps);
        skieur.Children.Add(baton);
        skieur.Children.Add(Rond(1, -22, 3.5, "#F2C9A0"));
        skieur.Children.Add(Chemin("M -2.6 -23 A 3.6 3.6 0 0 1 4.6 -23 Z", "#D93A3A"));
        skieur.Children.Add(Rond(1, -27.2, 1.3, "#FFFFFF"));
        var position = new TranslateTransform();
        var groupe = new TransformGroup();
        groupe.Children.Add(new ScaleTransform(k * 1.2, k * 1.2));
        groupe.Children.Add(position);
        skieur.RenderTransform = groupe;
        sc.Children.Add(skieur);

        const double duree = 14, t0 = 0.18, t1 = 0.86;
        sc.Animer(position, TranslateTransform.XProperty, Boucle(t0 * l, t1 * l, duree, TimeSpan.Zero));
        var trajetY = new DoubleAnimationUsingKeyFrames { Duration = TimeSpan.FromSeconds(duree), RepeatBehavior = RepeatBehavior.Forever };
        for (int i = 0; i <= 8; i++)
        {
            double f = i / 8.0, t = t0 + (t1 - t0) * f;
            double yt = (1 - t) * (1 - t) * ya + 2 * (1 - t) * t * yc + t * t * yb;
            trajetY.KeyFrames.Add(new LinearDoubleKeyFrame(yt, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(duree * f))));
        }
        sc.Animer(position, TranslateTransform.YProperty, trajetY);
        var fondu = new DoubleAnimationUsingKeyFrames { Duration = TimeSpan.FromSeconds(duree), RepeatBehavior = RepeatBehavior.Forever };
        fondu.KeyFrames.Add(new LinearDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.Zero)));
        fondu.KeyFrames.Add(new LinearDoubleKeyFrame(1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(1))));
        fondu.KeyFrames.Add(new LinearDoubleKeyFrame(1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(duree - 1))));
        fondu.KeyFrames.Add(new LinearDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(duree))));
        sc.Animer(skieur, UIElement.OpacityProperty, fondu);
    }

    private static void PaysagePaques(Scene sc, double l, double h, double k)
    {
        Collines(sc, l, h, 58 * k, 9 * k, 3, "#D9F0CF", 1, null);
        Collines(sc, l, h, 34 * k, 6 * k, 4, "#B5E0A6", 0, null);

        // Fleurs dans la prairie.
        var r = new Alea(17);
        for (double x = 14; x < l - 10; x += 46 * Math.Max(0.7, k))
        {
            double fx = x + r.Suivant() * 20, fy = h - (14 + r.Suivant() * 14) * k;
            var tige = new Forme.Line { X1 = fx, Y1 = fy, X2 = fx, Y2 = fy + 10 * k, Stroke = Pinceau("#5E9E57"), StrokeThickness = 1.2 };
            sc.Children.Add(tige);
            string couleur = Pastels[(int)(r.Suivant() * Pastels.Length) % Pastels.Length];
            for (int p = 0; p < 5; p++)
            {
                double a = p * 2 * Math.PI / 5;
                sc.Children.Add(Rond(fx + Math.Cos(a) * 3 * k, fy + Math.Sin(a) * 3 * k, 2.4 * k, couleur));
            }
            sc.Children.Add(Rond(fx, fy, 1.7 * k, "#F2B33D"));
        }

        // Œufs.
        double[] ox = { 0.84, 0.88, 0.91 };
        for (int i = 0; i < ox.Length; i++)
        {
            var oeuf = new Canvas();
            oeuf.Children.Add(new Forme.Ellipse { Width = 15, Height = 19, Fill = Pinceau(Pastels[(i * 2 + 1) % Pastels.Length]) });
            var motif = Chemin("M1.5 9 L4 7 L7.5 9.5 L11 7 L13.5 9", null);
            motif.Stroke = Brushes.White; motif.StrokeThickness = 1.4;
            oeuf.Children.Add(motif);
            var g = new TransformGroup();
            g.Children.Add(new RotateTransform(i == 1 ? 12 : -8, 7.5, 9.5));
            g.Children.Add(new ScaleTransform(k * 1.1, k * 1.1));
            oeuf.RenderTransform = g;
            Canvas.SetLeft(oeuf, ox[i] * l);
            Canvas.SetTop(oeuf, h - 28 * k);
            sc.Children.Add(oeuf);
        }

        // Papillons qui volettent.
        for (int i = 0; i < 2; i++)
        {
            var papillon = new Canvas();
            string aile = Pastels[(i * 3) % Pastels.Length];
            var ailes = new Canvas();
            var a1 = new Forme.Ellipse { Width = 8, Height = 11, Fill = Pinceau(aile) };
            Canvas.SetLeft(a1, -8); Canvas.SetTop(a1, -6);
            var a2 = new Forme.Ellipse { Width = 8, Height = 11, Fill = Pinceau(aile) };
            Canvas.SetLeft(a2, 0); Canvas.SetTop(a2, -6);
            ailes.Children.Add(a1);
            ailes.Children.Add(a2);
            var battement = new ScaleTransform(1, 1, 0, 0);
            ailes.RenderTransform = battement;
            papillon.Children.Add(ailes);
            papillon.Children.Add(Rect(-0.8, -5, 1.6, 10, 0.8, "#5B4636"));
            var vol = new TranslateTransform();
            var ondule = new TranslateTransform();
            var gv = new TransformGroup();
            gv.Children.Add(new ScaleTransform(k, k));
            gv.Children.Add(ondule);
            gv.Children.Add(vol);
            papillon.RenderTransform = gv;
            Canvas.SetTop(papillon, h - (70 + i * 22) * k);
            sc.Children.Add(papillon);
            double duree = 30 + i * 9;
            sc.Animer(vol, TranslateTransform.XProperty, Boucle(i == 0 ? -20 : l + 20, i == 0 ? l + 20 : -20, duree * Math.Max(1, l / 1100.0), TimeSpan.FromSeconds(-i * 11)));
            sc.Animer(ondule, TranslateTransform.YProperty, new DoubleAnimation(-10 * k, 10 * k, new Duration(TimeSpan.FromSeconds(1.7 + i * 0.4)))
            {
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
                EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
            });
            sc.Animer(battement, ScaleTransform.ScaleXProperty, new DoubleAnimation(1, 0.25, new Duration(TimeSpan.FromSeconds(0.18)))
            {
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
            });
        }

        // Lapin qui traverse en sautillant.
        var lapin = new Canvas();
        lapin.Children.Add(new Forme.Ellipse { Width = 22, Height = 15, Fill = Brushes.White, Stroke = Pinceau("#D8CCD6"), StrokeThickness = 1 });
        Canvas.SetLeft(lapin.Children[0], -11); Canvas.SetTop(lapin.Children[0], -15);
        var oreille1 = Oreille(6, -36, -8);
        var oreille2 = Oreille(11, -35, 12);
        lapin.Children.Add(oreille1);
        lapin.Children.Add(oreille2);
        var tete = Rond(11, -16, 6.5, "#FFFFFF");
        tete.Stroke = Pinceau("#D8CCD6"); tete.StrokeThickness = 1;
        lapin.Children.Add(tete);
        lapin.Children.Add(Rond(13.5, -17, 1, "#1F2937"));
        lapin.Children.Add(Rond(-11, -10, 3.4, "#FFFFFF"));
        var saut = new TranslateTransform();
        var course = new TranslateTransform();
        var gl = new TransformGroup();
        gl.Children.Add(new ScaleTransform(k, k));
        gl.Children.Add(saut);
        gl.Children.Add(course);
        lapin.RenderTransform = gl;
        Canvas.SetTop(lapin, h - 16 * k);
        sc.Children.Add(lapin);
        sc.Animer(course, TranslateTransform.XProperty, Boucle(-40, l + 40, 26 * Math.Max(1, l / 1100.0), TimeSpan.FromSeconds(-6)));
        sc.Animer(saut, TranslateTransform.YProperty, new DoubleAnimation(0, -13 * k, new Duration(TimeSpan.FromSeconds(0.42)))
        {
            AutoReverse = true,
            RepeatBehavior = RepeatBehavior.Forever,
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut },
        });
    }

    private static void PaysageEte(Scene sc, double l, double h, double k)
    {
        // Soleil dans le ciel, en haut à droite.
        var soleil = new Canvas();
        var rayons = new Forme.Ellipse
        {
            Width = 50,
            Height = 50,
            Stroke = Pinceau("#F6C33B"),
            StrokeThickness = 6,
            StrokeDashArray = new DoubleCollection { 0.6, 1.5 },
            Opacity = 0.85,
        };
        Canvas.SetLeft(rayons, -25); Canvas.SetTop(rayons, -25);
        var rotation = new RotateTransform(0, 25, 25);
        rayons.RenderTransform = rotation;
        soleil.Children.Add(rayons);
        soleil.Children.Add(Rond(0, 0, 15, "#F6C33B"));
        var gs = new TransformGroup();
        gs.Children.Add(new ScaleTransform(k, k));
        gs.Children.Add(new TranslateTransform(0.9 * l, h - 98 * k));
        soleil.RenderTransform = gs;
        sc.Children.Add(soleil);
        sc.Animer(rotation, RotateTransform.AngleProperty,
            new DoubleAnimation(0, 360, new Duration(TimeSpan.FromSeconds(40))) { RepeatBehavior = RepeatBehavior.Forever });

        // Mer (vagues qui défilent) puis sable.
        Vague(sc, l, h, 124 * k, 300, 4 * k, "#BFE6EA", 1, 30);
        Vague(sc, l, h, 104 * k, 220, 3 * k, "#9ED6DC", 0.85, 45);

        // Voilier qui dérive lentement sur l'eau.
        var bateau = new Canvas();
        bateau.Children.Add(Chemin("M -16 0 L 16 0 L 11 7 L -11 7 Z", "#1E2D4F"));
        var mat = new Forme.Line { X1 = 0, Y1 = 0, X2 = 0, Y2 = -30, Stroke = Pinceau("#1E2D4F"), StrokeThickness = 1.4 };
        bateau.Children.Add(mat);
        var voile = Chemin("M 1.5 -28 L 1.5 -2 L 17 -2 Z", "#FFFFFF");
        voile.Stroke = Pinceau("#C9D6E5"); voile.StrokeThickness = 1;
        bateau.Children.Add(voile);
        bateau.Children.Add(Chemin("M -1.5 -24 L -1.5 -2 L -12 -2 Z", "#F2B33D"));
        var tangage = new RotateTransform(0, 0, 4);
        var derive = new TranslateTransform();
        var gb = new TransformGroup();
        gb.Children.Add(tangage);
        gb.Children.Add(new ScaleTransform(k * 1.1, k * 1.1));
        gb.Children.Add(derive);
        bateau.RenderTransform = gb;
        Canvas.SetTop(bateau, h - 58 * k);
        sc.Children.Add(bateau);
        sc.Animer(derive, TranslateTransform.XProperty, Boucle(l + 40, -40, 70 * Math.Max(1, l / 1100.0), TimeSpan.FromSeconds(-25)));
        sc.Animer(tangage, RotateTransform.AngleProperty, new DoubleAnimation(-2.5, 2.5, new Duration(TimeSpan.FromSeconds(3)))
        {
            AutoReverse = true,
            RepeatBehavior = RepeatBehavior.Forever,
            EasingFunction = new SineEase { EasingMode = EasingMode.EaseInOut },
        });

        Collines(sc, l, h, 30 * k, 4 * k, 4, "#F3E3BF", 0, null);

        // Palmier.
        var palmier = new Canvas();
        var tronc = Chemin("M96 200 Q 92 140 104 74", null);
        tronc.Stroke = Pinceau("#B58A5A"); tronc.StrokeThickness = 9; tronc.StrokeStartLineCap = PenLineCap.Round; tronc.StrokeEndLineCap = PenLineCap.Round;
        palmier.Children.Add(tronc);
        foreach (var feuille in new[]
        {
            "M104 74 Q 70 50 22 70 Q 66 58 100 80 Z", "M104 74 Q 120 34 92 8 Q 128 34 108 78 Z",
            "M104 74 Q 146 46 184 62 Q 140 60 108 80 Z", "M104 74 Q 150 86 168 128 Q 138 96 104 82 Z",
            "M104 74 Q 64 86 46 128 Q 76 94 102 82 Z",
        })
            palmier.Children.Add(Chemin(feuille, "#4E9A5B"));
        double sp = 0.55 * k;
        var gp = new TransformGroup();
        gp.Children.Add(new ScaleTransform(sp, sp));
        gp.Children.Add(new TranslateTransform(0.07 * l - 96 * sp, h - 18 * k - 200 * sp));
        palmier.RenderTransform = gp;
        sc.Children.Add(palmier);

        // Parasol et serviette.
        double ux = 0.8 * l, uy = h - 20 * k;
        sc.Children.Add(Rect(ux - 2 * k, uy - 2 * k, 30 * k, 9 * k, 2 * k, "#A7CDF2"));
        sc.Children.Add(new Forme.Line { X1 = ux, Y1 = uy, X2 = ux, Y2 = uy - 34 * k, Stroke = Pinceau("#8A6A4A"), StrokeThickness = 1.6 });
        sc.Children.Add(Chemin(F($"M {ux - 24 * k} {uy - 30 * k} Q {ux} {uy - 54 * k} {ux + 24 * k} {uy - 30 * k} Z"), "#E89B2F"));
        sc.Children.Add(Chemin(F($"M {ux - 8 * k} {uy - 40 * k} Q {ux} {uy - 46 * k} {ux + 8 * k} {uy - 40 * k} L {ux} {uy - 30 * k} Z"), "#FFFFFF"));

        // Étoile de mer.
        var etoile = new StringBuilder();
        double ex = 0.58 * l, ey = h - 12 * k;
        for (int i = 0; i < 10; i++)
        {
            double a = -Math.PI / 2 + i * Math.PI / 5, rr = (i % 2 == 0 ? 6 : 2.6) * k;
            etoile.Append(F($"{(i == 0 ? "M" : " L")} {ex + Math.Cos(a) * rr} {ey + Math.Sin(a) * rr}"));
        }
        etoile.Append(" Z");
        sc.Children.Add(Chemin(etoile.ToString(), "#F08A5D"));
    }

    /// <summary>Collines arrondies sur toute la largeur, en bas de la scène.</summary>
    private static void Collines(Scene sc, double l, double h, double hauteur, double amplitude,
                                 int bosses, string couleur, int phase, string? contour)
    {
        double y0 = h - hauteur, pas = l / bosses;
        var d = new StringBuilder(F($"M 0 {h} L 0 {y0}"));
        for (int i = 0; i < bosses; i++)
        {
            double x0 = i * pas, sens = (i + phase) % 2 == 0 ? -1 : 1;
            d.Append(F($" Q {x0 + pas / 2} {y0 + sens * amplitude * 2} {x0 + pas} {y0}"));
        }
        d.Append(F($" L {l} {h} Z"));
        var colline = Chemin(d.ToString(), couleur);
        if (contour != null)
        {
            colline.Stroke = Pinceau(contour);
            colline.StrokeThickness = 1;
        }
        sc.Children.Add(colline);
    }

    /// <summary>Sapin de hauteur <paramref name="t"/>, pied en (x, yPied), pointes enneigées.</summary>
    private static Canvas Sapin(double x, double yPied, double t, string couleur, bool neige)
    {
        var c = new Canvas();
        c.Children.Add(Rect(-t * 0.06, -t * 0.14, t * 0.12, t * 0.14, 0, "#7A5A3A"));
        for (int i = 0; i < 3; i++)
        {
            double base_ = -t * 0.14 - i * t * 0.24, largeur = t * (0.62 - i * 0.14), sommet = base_ - t * 0.40;
            c.Children.Add(Chemin(F($"M {-largeur / 2} {base_} L {largeur / 2} {base_} L 0 {sommet} Z"), couleur));
            if (neige)
                c.Children.Add(Chemin(F($"M {-largeur * 0.17} {sommet + t * 0.12} L 0 {sommet} L {largeur * 0.17} {sommet + t * 0.12} Z"), "#FFFFFF"));
        }
        Canvas.SetLeft(c, x);
        Canvas.SetTop(c, yPied);
        return c;
    }

    /// <summary>Oreille de lapin (blanche, intérieur rose), inclinée de <paramref name="angle"/>.</summary>
    private static Canvas Oreille(double x, double y, double angle)
    {
        var o = new Canvas { Width = 10, Height = 26 };
        o.Children.Add(new Forme.Ellipse { Width = 10, Height = 26, Fill = Brushes.White, Stroke = Pinceau("#E5B8C8"), StrokeThickness = 1 });
        var dedans = new Forme.Ellipse { Width = 5, Height = 18, Fill = Pinceau("#F4A7B9") };
        Canvas.SetLeft(dedans, 2.5); Canvas.SetTop(dedans, 4);
        o.Children.Add(dedans);
        Canvas.SetLeft(o, x);
        Canvas.SetTop(o, y);
        o.RenderTransform = new RotateTransform(angle, 5, 26);
        return o;
    }

    private static void Mouettes(Scene sc, double l, double h)
    {
        int n = l > 800 ? 2 : 1;
        for (int i = 0; i < n; i++)
        {
            var oiseau = Chemin("M1 8 Q 7 1 13 8 Q 19 1 25 8", null);
            oiseau.Stroke = Pinceau("#1E2D4F");
            oiseau.StrokeThickness = 2;
            oiseau.StrokeStartLineCap = PenLineCap.Round;
            oiseau.StrokeEndLineCap = PenLineCap.Round;
            Canvas.SetTop(oiseau, i % 2 == 0 ? 6 : 20);

            var aile = new ScaleTransform(1, 1, 13, 8);
            var vol = new TranslateTransform();
            var groupe = new TransformGroup();
            groupe.Children.Add(aile);
            groupe.Children.Add(vol);
            oiseau.RenderTransform = groupe;
            sc.Children.Add(oiseau);

            double duree = 40 * l / 1180.0 + 8 + i * 7;
            sc.Animer(vol, TranslateTransform.XProperty, Boucle(l, -60, duree, TimeSpan.FromSeconds(-i * duree / 2)));
            sc.Animer(aile, ScaleTransform.ScaleYProperty, new DoubleAnimation(1, 0.55, new Duration(TimeSpan.FromSeconds(0.7)))
            {
                AutoReverse = true,
                RepeatBehavior = RepeatBehavior.Forever,
            });
        }
    }

    private static void Congere(Scene sc, double l)
    {
        var d = new StringBuilder("M 0 0 L 0 6");
        var r = new Alea(3);
        for (double x = 0; x < l; x += 50)
        {
            double creux = 8 + r.Suivant() * 5, bosse = 1 + r.Suivant() * 3;
            d.Append(F($" Q {x + 25} {creux} {x + 50} {bosse}"));
        }
        d.Append(F($" L {l + 50} 0 Z"));
        sc.Children.Add(Chemin(d.ToString(), "#FFFFFF"));
    }

    private static void Montagnes(Scene sc, double l, double h, bool sombre)
    {
        var groupe = new Canvas();
        if (sombre)
        {
            groupe.Children.Add(Chemin("M0 100 L90 38 L140 64 L230 10 L330 72 L400 30 L470 66 L560 22 L560 100 Z", "#2A3D63"));
            groupe.Children.Add(Chemin("M230 10 L262 30 L246 28 L232 38 L218 30 L204 28 Z", "#E6EEF8"));
            groupe.Children.Add(Chemin("M400 30 L424 45 L410 44 L398 51 L388 43 L378 46 Z", "#E6EEF8"));
            groupe.Children.Add(Chemin("M0 100 L60 70 L120 88 L200 52 L280 86 L360 58 L450 90 L520 64 L560 80 L560 100 Z", "#33507E"));
            groupe.Children.Add(Chemin("M200 52 L222 64 L210 63 L200 69 L190 62 L180 63 Z", "#F2F6FB"));
            groupe.Children.Add(Chemin("M360 58 L378 68 L368 68 L358 73 L350 67 L342 67 Z", "#F2F6FB"));
            // À l'échelle du bandeau, et jamais sur la moitié gauche (numéro de fiche).
            double kb = h / 100.0;
            groupe.RenderTransform = new ScaleTransform(kb, kb);
            Canvas.SetLeft(groupe, Math.Max(l * 0.35, l - 680 * kb));
            Canvas.SetTop(groupe, h - 100 * kb);
        }
        else
        {
            groupe.Children.Add(Chemin("M0 76 L60 34 L100 52 L170 8 L230 50 L270 28 L300 40 L300 76 Z", "#DCE7F3"));
            groupe.Children.Add(Chemin("M170 8 L194 23 L182 22 L171 29 L160 22 L148 23 Z", "#FFFFFF"));
            groupe.Children.Add(Chemin("M0 76 L50 56 L110 70 L180 44 L250 68 L300 54 L300 76 Z", "#C6D7EA"));
            double kb = h / 76.0;
            groupe.RenderTransform = new ScaleTransform(kb, kb);
            Canvas.SetLeft(groupe, Math.Max(l * 0.4, l - 300 * kb));
            Canvas.SetTop(groupe, h - 76 * kb);
        }
        sc.Children.Add(groupe);
    }

    private static void SoleilEtVagues(Scene sc, double l, double h, bool sombre)
    {
        // Soleil (rayons qui tournent lentement), à moitié sorti du bord haut.
        var soleil = new Canvas { Width = 120, Height = 120 };
        var rayons = new Forme.Ellipse
        {
            Width = 84,
            Height = 84,
            Stroke = Pinceau("#F6C33B"),
            StrokeThickness = 10,
            StrokeDashArray = new DoubleCollection { 0.6, 1.6 },
            Opacity = 0.85,
        };
        Canvas.SetLeft(rayons, 18); Canvas.SetTop(rayons, 18);
        var rotation = new RotateTransform(0, 42, 42);
        rayons.RenderTransform = rotation;
        soleil.Children.Add(rayons);
        var disque = Rond(60, 60, 27, "#F6C33B");
        soleil.Children.Add(disque);
        double ks = h / (sombre ? 100.0 : 95.0);
        soleil.RenderTransform = new ScaleTransform(ks, ks);
        Canvas.SetLeft(soleil, Math.Max(l * (sombre ? 0.42 : 0.45), l - (sombre ? 480 : 360) * ks));
        Canvas.SetTop(soleil, (sombre ? -46 : -54) * ks);
        sc.Children.Add(soleil);
        sc.Animer(rotation, RotateTransform.AngleProperty,
            new DoubleAnimation(0, 360, new Duration(TimeSpan.FromSeconds(40))) { RepeatBehavior = RepeatBehavior.Forever });

        // Vagues en bas, qui défilent sans couture (décalage d'une période).
        if (sombre)
        {
            Vague(sc, l, h, 26, 600, 5, "#7FD6D0", 0.35, 75);
            Vague(sc, l, h, 18, 400, 3.5, "#2BB3B1", 0.6, 130);
        }
        else
        {
            Vague(sc, l, h, 14, 400, 3, "#2BB3B1", 0.55, 120);
        }
    }

    private static void Vague(Scene sc, double l, double h, double hauteur, double periode,
                              double amplitude, string couleur, double opacite, double vitesse)
    {
        double milieu = hauteur / 2;
        var d = new StringBuilder(F($"M 0 {milieu} Q {periode / 4} {milieu - amplitude} {periode / 2} {milieu}"));
        for (double x = periode; x <= l + 2 * periode; x += periode / 2)
            d.Append(F($" T {x} {milieu}"));
        d.Append(F($" L {l + 2 * periode} {hauteur} L 0 {hauteur} Z"));
        var vague = Chemin(d.ToString(), couleur);
        vague.Opacity = opacite;
        Canvas.SetTop(vague, h - hauteur);
        var tt = new TranslateTransform();
        vague.RenderTransform = tt;
        sc.Children.Add(vague);
        sc.Animer(tt, TranslateTransform.XProperty, Boucle(0, -periode, periode / vitesse, TimeSpan.Zero));
    }

    // ------------------------------------------------------------------ outils

    /// <summary>Zone de dessin reconstruite à chaque changement de taille, dont les
    /// animations s'arrêtent quand elle quitte l'écran.</summary>
    private sealed class Scene : Canvas
    {
        private readonly Action<Scene, double, double> _construire;
        private readonly List<AnimationClock> _horloges = new();
        private double _largeur = -1, _hauteur = -1;

        public Scene(Action<Scene, double, double> construire)
        {
            _construire = construire;
            ClipToBounds = true;
            IsHitTestVisible = false;
            SizeChanged += (_, _) => Construire(false);
            Loaded += (_, _) => { if (Children.Count == 0) Construire(true); };
            Unloaded += (_, _) => Vider();
        }

        private void Construire(bool forcer)
        {
            double l = ActualWidth, h = ActualHeight;
            if (l < 1 || h < 1) return;
            if (!forcer && Math.Abs(l - _largeur) < 1 && Math.Abs(h - _hauteur) < 1) return;
            Vider();
            _largeur = l; _hauteur = h;
            _construire(this, l, h);
        }

        private void Vider()
        {
            foreach (var c in _horloges)
            {
                c.Controller?.Stop();
                Horloges.Remove(c);
            }
            _horloges.Clear();
            Children.Clear();
        }

        public void Animer(IAnimatable cible, DependencyProperty propriete, AnimationTimeline animation)
        {
            Timeline.SetDesiredFrameRate(animation, 30);
            var horloge = animation.CreateClock();
            cible.ApplyAnimationClock(propriete, horloge);
            _horloges.Add(horloge);
            Horloges.Add(horloge);
            if (_enPause) horloge.Controller?.Pause();
        }
    }

    /// <summary>Générateur pseudo-aléatoire déterministe (même décor à chaque lancement).</summary>
    private sealed class Alea
    {
        private long _s;
        public Alea(long graine) => _s = graine;
        public double Suivant() { _s = (_s * 9301 + 49297) % 233280; return _s / 233280.0; }
    }

    private static DoubleAnimation Boucle(double de, double a, double secondes, TimeSpan debut) =>
        new(de, a, new Duration(TimeSpan.FromSeconds(secondes)))
        {
            RepeatBehavior = RepeatBehavior.Forever,
            BeginTime = debut,
        };

    private static Forme.Path Chemin(string donnees, string? remplissage) => new()
    {
        Data = Geometry.Parse(donnees),
        Fill = remplissage is null ? null : Pinceau(remplissage),
    };

    private static Forme.Rectangle Rect(double x, double y, double l, double h, double rayon, string couleur)
    {
        var r = new Forme.Rectangle { Width = l, Height = h, RadiusX = rayon, RadiusY = rayon, Fill = Pinceau(couleur) };
        Canvas.SetLeft(r, x);
        Canvas.SetTop(r, y);
        return r;
    }

    private static Forme.Ellipse Rond(double cx, double cy, double rayon, string couleur)
    {
        var e = new Forme.Ellipse { Width = 2 * rayon, Height = 2 * rayon, Fill = Pinceau(couleur) };
        Canvas.SetLeft(e, cx - rayon);
        Canvas.SetTop(e, cy - rayon);
        return e;
    }

    private static SolidColorBrush Pinceau(string hex)
    {
        var b = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
        b.Freeze();
        return b;
    }

    /// <summary>Formate avec le point décimal (les géométries WPF l'exigent).</summary>
    private static string F(FormattableString s) => s.ToString(CultureInfo.InvariantCulture);
}
