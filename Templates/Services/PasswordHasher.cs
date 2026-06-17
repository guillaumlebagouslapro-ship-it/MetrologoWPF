using System;
using System.Security.Cryptography;

namespace Metrologo.Services
{
    /// <summary>
    /// Hash PBKDF2-HMAC-SHA256 avec sel aléatoire par compte.
    ///
    /// Format auto-décrit : <c>PBKDF2$&lt;iterations&gt;$&lt;sel-base64&gt;$&lt;hash-base64&gt;</c>
    /// Permet d'augmenter les itérations sans casser les anciens comptes, et de coexister
    /// avec Argon2/scrypt si migration future (préfixe distinct).
    ///
    /// 100 000 itérations = recommandation OWASP 2023 (~50 ms/hash, prohibitif en brute-force).
    /// </summary>
    public static class PasswordHasher
    {
        private const int IterationsParDefaut = 100_000;
        private const int TailleSelOctets = 16;     // 128 bits
        private const int TailleHashOctets = 32;    // 256 bits (sortie SHA256)
        private const string Algorithme = "PBKDF2";

        /// <summary>
        /// Génère un hash auto-décrit prêt à stocker dans <c>T_UTILISATEURS.PasswordHash</c>.
        /// </summary>
        public static string HashPassword(string password)
        {
            if (password == null) throw new ArgumentNullException(nameof(password));

            byte[] sel = RandomNumberGenerator.GetBytes(TailleSelOctets);
            byte[] hash = DeriverHash(password, sel, IterationsParDefaut);

            return $"{Algorithme}${IterationsParDefaut}${Convert.ToBase64String(sel)}${Convert.ToBase64String(hash)}";
        }

        /// <summary>
        /// Vérifie un mot de passe contre un hash stocké. Retourne <c>false</c> sur format
        /// invalide ou hash corrompu — pas d'exception (traité comme échec d'authentification).
        /// </summary>
        public static bool VerifyPassword(string password, string hashStocke)
        {
            if (string.IsNullOrEmpty(password) || string.IsNullOrEmpty(hashStocke)) return false;

            // Format : PBKDF2$iter$sel$hash
            string[] parties = hashStocke.Split('$');
            if (parties.Length != 4 || parties[0] != Algorithme) return false;
            if (!int.TryParse(parties[1], out int iterations) || iterations < 1) return false;

            byte[] sel, hashAttendu;
            try
            {
                sel = Convert.FromBase64String(parties[2]);
                hashAttendu = Convert.FromBase64String(parties[3]);
            }
            catch (FormatException) { return false; }

            byte[] hashCalcule = DeriverHash(password, sel, iterations);

            // Comparaison à temps constant — empêche les attaques par timing.
            return CryptographicOperations.FixedTimeEquals(hashCalcule, hashAttendu);
        }

        private static byte[] DeriverHash(string password, byte[] sel, int iterations)
        {
            using var pbkdf2 = new Rfc2898DeriveBytes(password, sel, iterations, HashAlgorithmName.SHA256);
            return pbkdf2.GetBytes(TailleHashOctets);
        }
    }
}
