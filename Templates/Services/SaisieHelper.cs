using System.Globalization;

namespace Metrologo.Services
{
    /// <summary>
    /// Regroupe au même endroit le parsing des champs numériques saisis par l'utilisateur.
    /// Plutôt que de recopier un <c>TryParse</c> dans chaque ViewModel, on gère ici une
    /// bonne fois pour toutes la virgule, le point et la culture invariante.
    /// </summary>
    public static class SaisieHelper
    {
        /// <summary>
        /// Met en forme une fréquence (Hz) pour l'affichage : on groupe la partie entière par
        /// milliers (« 10 000 000 ») sans rien perdre de la valeur exacte. On garde toute la
        /// précision du double, sans bruit flottant ni troncature, pour rester à la fois
        /// lisible et fidèle.
        /// </summary>
        public static string FormaterFrequence(double hz)
        {
            // On prend la représentation exacte la plus courte (sans bruit), en culture neutre.
            string brut = hz.ToString(CultureInfo.InvariantCulture);

            // Si on tombe sur de la notation scientifique (valeurs extrêmes), repli simple groupé.
            if (brut.IndexOf('E') >= 0 || brut.IndexOf('e') >= 0)
                return hz.ToString("#,##0.######", CultureInfo.GetCultureInfo("fr-FR"));

            bool negatif = brut.StartsWith("-", System.StringComparison.Ordinal);
            if (negatif) brut = brut.Substring(1);

            string partieEntiere = brut;
            string partieFrac = string.Empty;
            int pt = brut.IndexOf('.');
            if (pt >= 0)
            {
                partieEntiere = brut.Substring(0, pt);
                partieFrac = brut.Substring(pt + 1);
            }

            var frFr = CultureInfo.GetCultureInfo("fr-FR");
            string sep = frFr.NumberFormat.NumberGroupSeparator;   // séparateur de milliers fr-FR

            // Partie entière : on groupe par paquets de 3 en partant de la virgule vers la gauche.
            if (long.TryParse(partieEntiere, NumberStyles.None, CultureInfo.InvariantCulture, out long ent))
                partieEntiere = ent.ToString("#,##0", frFr);

            // Et on groupe pareil la partie décimale, mais en allant de la virgule vers la droite.
            if (partieFrac.Length > 3)
            {
                var sb = new System.Text.StringBuilder(partieFrac.Length + partieFrac.Length / 3);
                for (int i = 0; i < partieFrac.Length; i++)
                {
                    if (i > 0 && i % 3 == 0) sb.Append(sep);
                    sb.Append(partieFrac[i]);
                }
                partieFrac = sb.ToString();
            }

            return (negatif ? "-" : string.Empty)
                 + partieEntiere
                 + (partieFrac.Length > 0 ? "," + partieFrac : string.Empty);
        }

        /// <summary>
        /// Convertit un texte saisi en <see cref="double"/>. On accepte aussi bien la virgule
        /// que le point comme séparateur décimal, on rogne les espaces de bord, et on s'appuie
        /// sur <see cref="CultureInfo.InvariantCulture"/> pour que le résultat ne dépende pas
        /// de la locale de la machine.
        /// </summary>
        public static bool TryParseDouble(string? s, out double value)
        {
            return double.TryParse(
                (s ?? string.Empty).Trim().Replace(',', '.'),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out value);
        }

        /// <summary>
        /// Même chose que <see cref="TryParseDouble"/>, mais on refuse tout ce qui n'est pas
        /// strictement positif (&gt; 0). Pratique pour les fréquences, les durées, etc.
        /// </summary>
        public static bool TryParsePositiveDouble(string? s, out double value)
        {
            return TryParseDouble(s, out value) && value > 0;
        }

        /// <summary>
        /// Même chose que <see cref="TryParseDouble"/>, mais ici on tolère le zéro et on
        /// rejette seulement le négatif (&gt;= 0). Pratique pour les résolutions, les
        /// incertitudes, etc.
        /// </summary>
        public static bool TryParseNonNegativeDouble(string? s, out double value)
        {
            return TryParseDouble(s, out value) && value >= 0;
        }

        /// <summary>
        /// Dit si <paramref name="text"/> ressemble à un début de nombre <em>plausible</em>
        /// pendant la frappe (chaîne vide, "-", "1.", "1,2", "1e-", "3E+04"...). C'est ce que
        /// le behavior de filtrage clavier utilise pour laisser passer les états intermédiaires
        /// tout en bloquant les lettres parasites.
        /// </summary>
        public static bool IsPlausibleNumericInput(string? text, bool allowNegative, bool allowDecimal)
        {
            if (string.IsNullOrEmpty(text))
                return true;

            bool sawDigit = false, sawSep = false, sawExp = false, sawExpSign = false, sawExpDigit = false;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                if (c >= '0' && c <= '9')
                {
                    if (sawExp) sawExpDigit = true; else sawDigit = true;
                    continue;
                }

                if (c == '-' || c == '+')
                {
                    // Un signe n'est valide qu'en tête de chaîne ou juste après un 'e'/'E'.
                    bool afterExp = sawExp && !sawExpSign && !sawExpDigit;
                    bool atStart = i == 0;
                    if (c == '-' && atStart && !allowNegative) return false;
                    if (!atStart && !afterExp) return false;
                    if (afterExp) sawExpSign = true;
                    continue;
                }

                if (c == '.' || c == ',')
                {
                    if (!allowDecimal || sawSep || sawExp) return false;
                    sawSep = true;
                    continue;
                }

                if (c == 'e' || c == 'E')
                {
                    if (sawExp || !sawDigit) return false;
                    sawExp = true;
                    continue;
                }

                return false; // n'importe quel autre caractère disqualifie la saisie
            }

            return true;
        }
    }
}
