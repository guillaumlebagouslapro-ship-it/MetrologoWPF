using System.Security.Cryptography;
using System.Text;

namespace Metrologo.Services
{
    /// <summary>
    /// Génère un mot de passe initial pour un nouvel utilisateur.
    /// 12 caractères, alphabet « lisible » (pas de O/0, I/l/1) pour éviter les
    /// confusions à la dictée orale. Garantit au moins 1 majuscule, 1 minuscule,
    /// 1 chiffre, 1 symbole — couvre les politiques de complexité standard.
    ///
    /// Génération via <see cref="RandomNumberGenerator"/> : crypto-secure, sans
    /// biais modulo. L'admin communique le mot de passe à l'utilisateur, qui
    /// devra le changer au premier login (TODO : workflow de changement).
    /// </summary>
    public static class MotDePasseGenerator
    {
        // Alphabets sans caractères ambigus : pas de O/0, I/l/1, S/5, B/8.
        private const string Majuscules = "ACDEFGHJKLMNPQRTUVWXYZ";
        private const string Minuscules = "abcdefghijkmnpqrtuvwxyz";
        private const string Chiffres   = "23456789";
        private const string Symboles   = "@#$%&*+=?";

        public static string Generer(int longueur = 12)
        {
            if (longueur < 4) longueur = 4;

            var sortie = new char[longueur];

            // 4 premiers caractères : un de chaque catégorie (garantit la complexité).
            sortie[0] = Pick(Majuscules);
            sortie[1] = Pick(Minuscules);
            sortie[2] = Pick(Chiffres);
            sortie[3] = Pick(Symboles);

            string tousAlphabets = Majuscules + Minuscules + Chiffres + Symboles;
            for (int i = 4; i < longueur; i++)
                sortie[i] = Pick(tousAlphabets);

            // Mélange Fisher-Yates pour que les 4 caractères « catégoriels » ne soient
            // pas systématiquement aux 4 premières positions (sinon prédictible).
            for (int i = longueur - 1; i > 0; i--)
            {
                int j = RandomNumberGenerator.GetInt32(i + 1);
                (sortie[i], sortie[j]) = (sortie[j], sortie[i]);
            }

            return new string(sortie);
        }

        private static char Pick(string alphabet)
            => alphabet[RandomNumberGenerator.GetInt32(alphabet.Length)];
    }
}
