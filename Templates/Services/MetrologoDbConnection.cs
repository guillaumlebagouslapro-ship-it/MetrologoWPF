using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Data.SqlClient;

namespace Metrologo.Services
{
    /// <summary>
    /// Point unique pour obtenir une connexion SQL Server vers la base centralisée Metrologo.
    ///
    /// La connection string est résolue dans cet ordre :
    ///   1. Variable d'environnement <c>METROLOGO_DB_CONNSTRING</c> (override déploiement).
    ///   2. Fichier <c>%LocalAppData%\Metrologo\db.credentials</c> (login SQL chiffré DPAPI).
    ///   3. Fichier <c>%LocalAppData%\Metrologo\database.connstring.txt</c> (legacy / debug).
    ///   4. Valeur par défaut (auth Windows locale, fallback dev).
    ///
    /// Le chemin (2) est le mode prod : un login SQL dédié (<c>metrologo_app</c>) avec
    /// permissions limitées à la base Metrologo, mot de passe chiffré via DPAPI lié au
    /// profil Windows courant. Conséquence : un autre utilisateur Windows ne peut pas
    /// lire le credentials même s'il a accès au fichier (DPAPI utilise une clé dérivée
    /// du compte). Aucun mot de passe SQL n'est jamais stocké en clair sur disque.
    /// </summary>
    public static class MetrologoDbConnection
    {
        private const string ConnStringParDefaut =
            @"Server=localhost\SQLEXPRESS;Database=Metrologo;Trusted_Connection=True;TrustServerCertificate=True;Encrypt=False;Connect Timeout=5";

        private const string NomVariableEnv = "METROLOGO_DB_CONNSTRING";
        private const string NomFichierConfig = "database.connstring.txt";
        private const string NomFichierCredentials = "db.credentials";

        private static string? _connStringEnCache;

        public static string ConnectionString
        {
            get
            {
                if (_connStringEnCache != null) return _connStringEnCache;

                // 1. Variable d'env (priorité max — utilisée pour CI/CD, override admin)
                string? viaEnv = Environment.GetEnvironmentVariable(NomVariableEnv);
                if (!string.IsNullOrWhiteSpace(viaEnv))
                {
                    _connStringEnCache = viaEnv;
                    return _connStringEnCache;
                }

                string dossier = DossierConfig();

                // 2. Login SQL chiffré DPAPI (mode prod)
                string cheminCreds = Path.Combine(dossier, NomFichierCredentials);
                if (File.Exists(cheminCreds))
                {
                    try
                    {
                        var creds = LireCredentialsDpapi(cheminCreds);
                        if (creds != null)
                        {
                            _connStringEnCache = ConstruireConnString(creds.Value.user, creds.Value.password);
                            return _connStringEnCache;
                        }
                    }
                    catch (Exception)
                    {
                        // Credentials corrompus ou DPAPI illisible (autre profil) — on tombe
                        // sur le mode auth Windows, l'utilisateur verra l'erreur au login.
                    }
                }

                // 3. Connection string brute en clair (legacy, débuggage)
                string cheminConn = Path.Combine(dossier, NomFichierConfig);
                if (File.Exists(cheminConn))
                {
                    try
                    {
                        string contenu = File.ReadAllText(cheminConn).Trim();
                        if (!string.IsNullOrWhiteSpace(contenu))
                        {
                            _connStringEnCache = contenu;
                            return _connStringEnCache;
                        }
                    }
                    catch (Exception) { /* fallback */ }
                }

                // 4. Défaut (dev local, auth Windows)
                _connStringEnCache = ConnStringParDefaut;
                return _connStringEnCache;
            }
        }

        public static SqlConnection Ouvrir()
        {
            var c = new SqlConnection(ConnectionString);
            c.Open();
            return c;
        }

        public static void ReinitialiserCache() => _connStringEnCache = null;

        // -------------------------------------------------------------------------
        // Stockage credentials chiffré (DPAPI)
        // -------------------------------------------------------------------------

        /// <summary>
        /// Chiffre les credentials SQL avec DPAPI (clé liée au profil Windows courant)
        /// et les écrit dans <c>%LocalAppData%\Metrologo\db.credentials</c>. À appeler
        /// une seule fois lors du provisionnement du poste.
        /// </summary>
        public static void EcrireCredentialsDpapi(string user, string password)
        {
            string dossier = DossierConfig();
            Directory.CreateDirectory(dossier);

            // Format en clair avant chiffrement : "user\npassword"
            byte[] enClair = Encoding.UTF8.GetBytes($"{user}\n{password}");
            byte[] chiffre = ProtectedData.Protect(enClair, optionalEntropy: null,
                scope: DataProtectionScope.CurrentUser);

            string chemin = Path.Combine(dossier, NomFichierCredentials);
            File.WriteAllBytes(chemin, chiffre);
            ReinitialiserCache();
        }

        private static (string user, string password)? LireCredentialsDpapi(string chemin)
        {
            byte[] chiffre = File.ReadAllBytes(chemin);
            byte[] enClair = ProtectedData.Unprotect(chiffre, optionalEntropy: null,
                scope: DataProtectionScope.CurrentUser);
            string s = Encoding.UTF8.GetString(enClair);
            int sep = s.IndexOf('\n');
            if (sep <= 0) return null;
            return (s[..sep], s[(sep + 1)..]);
        }

        private static string ConstruireConnString(string user, string password)
            => $"Server=localhost\\SQLEXPRESS;Database=Metrologo;User Id={user};Password={password};TrustServerCertificate=True;Encrypt=False;Connect Timeout=5";

        private static string DossierConfig() => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Metrologo");
    }
}
