using Metrologo.Services.Journal;
using Microsoft.Data.SqlClient;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services
{
    /// <summary>
    /// Accès à la base ASERi (SQL Server <c>SVR-OR</c>, base <c>SIA</c>) pour vérifier qu'un N° de
    /// fiche d'intervention existe. Portage du flux Delphi <c>F_Configuration.pas:104-127</c>
    /// (la validation de longueur « 8 caractères » de l'ancien flux est faite par l'appelant).
    /// <para/>La chaîne de connexion est codée en dur : ASERi est une base interne au LAN, partagée
    /// par plusieurs applications métier. Pour la sécuriser, reprendre le pattern DPAPI de
    /// <c>db.credentials</c>.
    /// </summary>
    public static class AseriService
    {
        // Reprise de la connexion OLE DB du Metrologo.exe Delphi (Provider=SQLNCLI.1;...) pour
        // Microsoft.Data.SqlClient. TrustServerCertificate=true : les vieux SQL Server
        // (2008/2012) n'ont pas de certificat valide.
        private const string ConnectionString =
            "Server=SVR-OR;Database=SIA;User Id=russe;Password=cia;"
          + "TrustServerCertificate=true;Encrypt=false;"
          + "Connect Timeout=5";

        /// <summary>Délai maximal d'interrogation d'ASERi, en secondes.</summary>
        private const int TimeoutSecondes = 5;

        /// <summary>
        /// Le N° de FI figure-t-il dans <c>SIA..tAffaire</c> ? Renvoie <c>null</c> si ASERi est
        /// injoignable : à l'appelant de décider d'accepter ou de refuser.
        /// </summary>
        public static async Task<bool?> FiExisteAsync(string numFITrim)
        {
            if (string.IsNullOrWhiteSpace(numFITrim)) return false;

            // ASERi stocke les n° FI avec / (D9/00000) ; l'UI utilise _ (D9_00000) pour rester
            // valide dans les chemins de fichiers. Cf. F_Configuration.pas:115.
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
                return null;   // échec de connexion : l'appelant tranche
            }
        }
    }
}
