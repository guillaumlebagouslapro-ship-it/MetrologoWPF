using Metrologo.Services.Journal;
using Microsoft.Data.SqlClient;
using System;
using System.Threading.Tasks;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services
{
    /// <summary>
    /// Donne accès à la base ASERi (SQL Server <c>SVR-OR</c>, base <c>SIA</c>) pour vérifier
    /// qu'un N° de fiche d'intervention existe bien.
    ///
    /// <para/>C'est le portage à l'identique du flux Delphi <c>F_Configuration.pas:104-127</c> :
    /// <list type="number">
    ///   <item>On trim le texte saisi.</item>
    ///   <item>On exige une longueur de pile 8 caractères, sinon « N° de FI incorrect ».</item>
    ///   <item>On remplace <c>_</c> par <c>/</c> avant la requête, car la BDD stocke avec /.</item>
    ///   <item>On lance <c>SELECT TOP 1 AffID FROM SIA..tAffaire WHERE AffNoFI=@fi</c>.</item>
    ///   <item>Sans résultat → « La FI n° X n'existe pas dans ASERi ».</item>
    /// </list>
    ///
    /// <para/>La chaîne de connexion est codée en dur ici. ASERi est une BDD interne au réseau,
    /// partagée par plusieurs applis métier : au-delà du LAN, ça ne révèle rien de sensible.
    /// Si on doit un jour la sécuriser (chiffrement DPAPI), on reprendra le pattern déjà utilisé
    /// pour <c>db.credentials</c> du Metrologo SQL principal.
    /// </summary>
    public static class AseriService
    {
        // On a reconstitué cette chaîne à partir de la connexion OLE DB embarquée dans
        // l'ancien Metrologo.exe Delphi (Provider=SQLNCLI.1;...), en la transposant pour
        // Microsoft.Data.SqlClient (.NET).
        // TrustServerCertificate=true est incontournable avec les vieux SQL Server
        // (2008/2012) qui n'ont pas de certificat valide installé.
        private const string ConnectionString =
            "Server=SVR-OR;Database=SIA;User Id=russe;Password=cia;"
          + "TrustServerCertificate=true;Encrypt=false;"
          + "Connect Timeout=5";

        /// <summary>Le temps qu'on s'accorde pour interroger ASERi : passé 5 s sans réponse, on considère que c'est KO.</summary>
        private const int TimeoutSecondes = 5;

        /// <summary>
        /// Regarde si un N° de FI figure dans la table <c>SIA..tAffaire</c>. Renvoie <c>true</c>
        /// quand on le trouve, <c>false</c> sinon. Et si la connexion à ASERi échoue (réseau
        /// coupé, serveur injoignable), on renvoie <c>null</c> : libre à l'appelant de laisser
        /// passer quand même ou de refuser.
        /// </summary>
        public static async Task<bool?> FiExisteAsync(string numFITrim)
        {
            if (string.IsNullOrWhiteSpace(numFITrim)) return false;

            // On repasse les _ en / : ASERi stocke les n° FI avec / (ex. D9/00000), alors
            // que côté UI on saisit avec _ (D9_00000) pour ne pas se créer d'ennuis dans les
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
                return null;   // ici on signale juste l'échec de connexion, c'est l'appelant qui tranche
            }
        }
    }
}
