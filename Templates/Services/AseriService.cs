using Metrologo.Services.Journal;
using Microsoft.Data.SqlClient;
using System;
using System.Threading.Tasks;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services
{
    /// <summary>
    /// Accès à la base ASERi (SQL Server <c>SVR-OR</c>, base <c>SIA</c>) pour
    /// valider qu'un N° de fiche d'intervention existe bien.
    ///
    /// <para/>Portage strict du flux Delphi <c>F_Configuration.pas:104-127</c> :
    /// <list type="number">
    ///   <item>Trim du texte saisi.</item>
    ///   <item>Longueur exacte = 8 caractères (sinon « N° de FI incorrect »).</item>
    ///   <item>Remplacement <c>_</c> → <c>/</c> avant requête (la BDD stocke avec /).</item>
    ///   <item>Requête <c>SELECT TOP 1 AffID FROM SIA..tAffaire WHERE AffNoFI=@fi</c></item>
    ///   <item>Si pas de résultat → « La FI n° X n'existe pas dans ASERi ».</item>
    /// </list>
    ///
    /// <para/>La chaîne de connexion est en dur ici (ASERi est une BDD interne au
    /// réseau, utilisée par plusieurs apps métier — ce n'est pas un secret au-delà
    /// du LAN). Si besoin de la durcir (chiffrement DPAPI), on suivra le pattern
    /// déjà en place pour <c>db.credentials</c> du Metrologo SQL principal.
    /// </summary>
    public static class AseriService
    {
        // Chaîne reconstituée à partir de la connection OLE DB embarquée dans le
        // Metrologo.exe Delphi historique (Provider=SQLNCLI.1;...), transposée
        // pour Microsoft.Data.SqlClient (.NET).
        // TrustServerCertificate=true → indispensable pour les SQL Server anciens
        // (2008/2012) sans certificat valide installé.
        private const string ConnectionString =
            "Server=SVR-OR;Database=SIA;User Id=russe;Password=cia;"
          + "TrustServerCertificate=true;Encrypt=false;"
          + "Connect Timeout=5";

        /// <summary>Délai max pour interroger ASERi. Si la BDD ne répond pas en 5 s, on considère KO.</summary>
        private const int TimeoutSecondes = 5;

        /// <summary>
        /// Vérifie qu'un N° de FI existe dans la table <c>SIA..tAffaire</c>. Retourne
        /// <c>true</c> si trouvé, <c>false</c> sinon. En cas d'erreur de connexion à
        /// ASERi (réseau down, serveur inaccessible), retourne <c>null</c> pour que
        /// l'appelant puisse décider d'autoriser quand même ou de refuser.
        /// </summary>
        public static async Task<bool?> FiExisteAsync(string numFITrim)
        {
            if (string.IsNullOrWhiteSpace(numFITrim)) return false;

            // Conversion _ → / : la BDD ASERi stocke les n° FI avec / (ex. D9/00000)
            // alors que côté UI on saisit avec _ (D9_00000) pour éviter les pbs de
            // chemins de fichiers. Cf. F_Configuration.pas:115.
            string sFI = numFITrim.Replace('_', '/');

            const string sql = "SELECT TOP 1 AffID FROM SIA..tAffaire WHERE AffNoFI=@fi";

            try
            {
                using var conn = new SqlConnection(ConnectionString);
                await conn.OpenAsync();

                using var cmd = new SqlCommand(sql, conn)
                {
                    CommandTimeout = TimeoutSecondes
                };
                cmd.Parameters.Add(new SqlParameter("@fi", System.Data.SqlDbType.NVarChar, 50) { Value = sFI });

                var result = await cmd.ExecuteScalarAsync();
                bool existe = result != null && result != DBNull.Value;

                if (existe)
                {
                    JournalLog.Info(CategorieLog.Configuration, "ASERI_FI_OK",
                        $"FI vérifiée dans ASERi : {sFI} → AffID={result}");
                }
                else
                {
                    JournalLog.Warn(CategorieLog.Configuration, "ASERI_FI_INTROUVABLE",
                        $"FI saisie introuvable dans ASERi : {sFI}");
                }
                return existe;
            }
            catch (Exception ex)
            {
                JournalLog.Erreur(CategorieLog.Configuration, "ASERI_CONNEXION_KO",
                    $"Impossible de joindre ASERi (SVR-OR/SIA) pour vérifier la FI {sFI} : {ex.Message}");
                return null;   // signal d'erreur connexion — l'appelant décide
            }
        }
    }
}
