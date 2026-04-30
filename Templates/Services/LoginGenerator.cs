using System.Globalization;
using System.Linq;
using System.Text;

namespace Metrologo.Services
{
    /// <summary>
    /// Génère un login canonique au format <c>prenom.nom</c> à partir des saisies admin.
    /// Normalise : retire les accents, met en minuscules, ne garde que [a-z0-9.-],
    /// remplace les espaces par des tirets pour les noms composés (« De Pas » → « de-pas »).
    /// La détection de collision est gérée par <see cref="UtilisateursService"/> qui
    /// ajoute un suffixe numérique si nécessaire.
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

            // Décomposition Unicode pour retirer les diacritiques (é → e, ç → c, etc.)
            string decompose = s.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder(decompose.Length);
            foreach (char c in decompose)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) == UnicodeCategory.NonSpacingMark) continue;
                sb.Append(c);
            }
            string sansAccents = sb.ToString().Normalize(NormalizationForm.FormC).ToLowerInvariant();

            // Espaces → tirets, on retire tout ce qui n'est pas [a-z0-9.-]
            var sortie = new StringBuilder(sansAccents.Length);
            foreach (char c in sansAccents)
            {
                if (c == ' ' || c == '_') sortie.Append('-');
                else if ((c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '.' || c == '-')
                    sortie.Append(c);
                // tout le reste : ignoré (apostrophes, ponctuation, etc.)
            }

            // Trim des séparateurs en bordure et compaction des doublons (--)
            string r = sortie.ToString().Trim('-', '.');
            while (r.Contains("--")) r = r.Replace("--", "-");
            return r;
        }
    }
}
