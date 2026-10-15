using System.Globalization;
using System.Text;

namespace Metrologo.Services
{
    /// <summary>
    /// Génère un login <c>prenom.nom</c> : sans accents, en minuscules, restreint à [a-z0-9.-],
    /// espaces en tirets (« De Pas » → « de-pas »). Les collisions sont gérées par
    /// <see cref="ComptesLocauxService"/> (suffixe numérique).
    /// </summary>
    public static class LoginGenerator
    {
        public static string GenererBase(string prenom, string nom)
        {
            string p = Normaliser(prenom);
            string n = Normaliser(nom);
            if (p.Length == 0 && n.Length == 0) return "utilisateur";
            if (p.Length == 0) return n;
            if (n.Length == 0) return p;
            return $"{p}.{n}";
        }

        private static string Normaliser(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;

            // Décomposition Unicode pour retirer les diacritiques (é → e, ç → c…)
            string decompose = s.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(decompose.Length);
            foreach (char c in decompose)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) == UnicodeCategory.NonSpacingMark) continue;
                sb.Append(c);
            }
            string sansAccents = sb.ToString().Normalize(NormalizationForm.FormC).ToLowerInvariant();

            // Espaces → tirets ; hors [a-z0-9.-] supprimé
            var sortie = new StringBuilder(sansAccents.Length);
            foreach (char c in sansAccents)
            {
                if (c == ' ' || c == '_') sortie.Append('-');
                else if ((c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '.' || c == '-')
                    sortie.Append(c);
                // tout le reste : ignoré (apostrophes, ponctuation, etc.)
            }

            // Trim des séparateurs et compaction des doublons (--)
            string r = sortie.ToString().Trim('-', '.');
            while (r.Contains("--")) r = r.Replace("--", "-");
            return r;
        }
    }
}
