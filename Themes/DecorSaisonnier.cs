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
        double xLapin = l - (sombre ? 470 : 340);

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
            Canvas.SetLeft(groupe, l - 680);
            Canvas.SetTop(groupe, h - 100);
        }
        else
        {
            groupe.Children.Add(Chemin("M0 76 L60 34 L100 52 L170 8 L230 50 L270 28 L300 40 L300 76 Z", "#DCE7F3"));
            groupe.Children.Add(Chemin("M170 8 L194 23 L182 22 L171 29 L160 22 L148 23 Z", "#FFFFFF"));
            groupe.Children.Add(Chemin("M0 76 L50 56 L110 70 L180 44 L250 68 L300 54 L300 76 Z", "#C6D7EA"));
            Canvas.SetLeft(groupe, l - 300);
            Canvas.SetTop(groupe, h - 76);
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
        Canvas.SetLeft(soleil, l - (sombre ? 480 : 360));
        Canvas.SetTop(soleil, sombre ? -46 : -54);
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
