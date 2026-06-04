using System.Globalization;

namespace Metrologo.Services
{
    /// <summary>
    /// Helpers centralisés de parsing des champs de saisie numériques.
    /// Remplace les implémentations <c>TryParse</c> dupliquées dans chaque
    /// ViewModel : gestion virgule/point et culture invariante en un seul endroit.
    /// </summary>
    public static class SaisieHelper
    {
        /// <summary>
        /// Formate une fréquence (Hz) pour AFFICHAGE : partie entière groupée par milliers
        /// (« 10 000 000 ») tout en conservant la valeur EXACTE (toute la précision du double,
        /// sans bruit flottant ni troncature). Combine lisibilité et exactitude.
        /// </summary>
        public static string FormaterFrequence(double hz)
        {
            // Représentation exacte la plus courte (pas de bruit), culture neutre.
            string brut = hz.ToString(CultureInfo.InvariantCulture);

            // Cas notation scientifique (valeurs extrêmes) : repli simple groupé.
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

            // Groupe la partie entière par paquets de 3 (depuis la virgule vers la gauche).
            if (long.TryParse(partieEntiere, NumberStyles.None, CultureInfo.InvariantCulture, out long ent))
                partieEntiere = ent.ToString("#,##0", frFr);

            // Groupe AUSSI la partie décimale par paquets de 3 (depuis la virgule vers la droite).
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
        /// Parse un texte saisi en <see cref="double"/>. Accepte la virgule ou le
        /// point comme séparateur décimal, ignore les espaces de bord, et utilise
        /// <see cref="CultureInfo.InvariantCulture"/> pour un résultat indépendant
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
        /// Comme <see cref="TryParseDouble"/> mais exige une valeur strictement
        /// positive (&gt; 0). Pratique pour les fréquences, durées, etc.
        /// </summary>
        public static bool TryParsePositiveDouble(string? s, out double value)
        {
            return TryParseDouble(s, out value) && value > 0;
        }

        /// <summary>
        /// Comme <see cref="TryParseDouble"/> mais exige une valeur positive ou
        /// nulle (&gt;= 0). Pratique pour les résolutions, incertitudes, etc.
        /// </summary>
        public static bool TryParseNonNegativeDouble(string? s, out double value)
        {
            return TryParseDouble(s, out value) && value >= 0;
        }

        /// <summary>
        /// Indique si <paramref name="text"/> est un préfixe numérique <em>plausible</em>
        /// pendant la frappe (chaîne vide, "-", "1.", "1,2", "1e-", "3E+04"...).
        /// Utilisé par le behavior de filtrage clavier pour autoriser les états
        /// intermédiaires sans laisser passer de lettres parasites.
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
                    // Signe accepté en tête, ou juste après un 'e'/'E'.
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

                return false; // tout autre caractère => invalide
            }

            return true;
        }
    }
}
