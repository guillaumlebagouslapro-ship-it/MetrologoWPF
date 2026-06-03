using System;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using Metrologo.Services.Journal;

namespace Metrologo.Services.Besancon
{
    /// <summary>
    /// Récupération du fichier de l'observatoire de Besançon sur le FTP (équivalent de la
    /// récupération Indy <c>ftpe2m</c> du legacy). Utilise <c>FtpWebRequest</c> (intégré .NET,
    /// pas de dépendance externe) ; gère le FTPS explicite via <see cref="BesanconConfig.FtpSsl"/>.
    /// </summary>
    public static class BesanconFtpService
    {
        /// <summary>Télécharge le fichier distant et retourne son contenu texte (null si échec/non configuré).</summary>
        public static async Task<string?> TelechargerAsync(BesanconConfig cfg)
        {
            if (string.IsNullOrWhiteSpace(cfg.FtpHote) || string.IsNullOrWhiteSpace(cfg.FtpUtilisateur))
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_FTP_CONFIG",
                    "FTP Besançon non configuré (hôte ou identifiant manquant) — téléchargement ignoré. "
                  + $"Renseigne {BesanconConfig.Chemin}.");
                return null;
            }

            string url = $"ftp://{cfg.FtpHote}:{cfg.FtpPort}/{cfg.FichierDistant}";
            try
            {
                return await Task.Run(() =>
                {
#pragma warning disable SYSLIB0014   // FtpWebRequest obsolète mais suffisant pour un GET simple, sans dépendance
                    var req = (FtpWebRequest)WebRequest.Create(url);
#pragma warning restore SYSLIB0014
                    req.Method = WebRequestMethods.Ftp.DownloadFile;
                    req.Credentials = new NetworkCredential(cfg.FtpUtilisateur, cfg.FtpMotDePasse);
                    req.UseBinary = true;
                    req.UsePassive = true;
                    req.EnableSsl = cfg.FtpSsl;
                    req.Timeout = 30000;

                    using var resp = (FtpWebResponse)req.GetResponse();
                    using var stream = resp.GetResponseStream();
                    using var reader = new StreamReader(stream!);
                    return reader.ReadToEnd();
                });
            }
            catch (Exception ex)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_FTP_KO",
                    $"Téléchargement FTP Besançon échoué ({url}) : {ex.Message}");
                return null;
            }
        }

        /// <summary>Supprime le fichier distant après récupération (comportement legacy, optionnel).</summary>
        public static async Task SupprimerDistantAsync(BesanconConfig cfg)
        {
            string url = $"ftp://{cfg.FtpHote}:{cfg.FtpPort}/{cfg.FichierDistant}";
            try
            {
                await Task.Run(() =>
                {
#pragma warning disable SYSLIB0014
                    var req = (FtpWebRequest)WebRequest.Create(url);
#pragma warning restore SYSLIB0014
                    req.Method = WebRequestMethods.Ftp.DeleteFile;
                    req.Credentials = new NetworkCredential(cfg.FtpUtilisateur, cfg.FtpMotDePasse);
                    req.EnableSsl = cfg.FtpSsl;
                    req.Timeout = 30000;
                    using var resp = (FtpWebResponse)req.GetResponse();
                });
            }
            catch (Exception ex)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_FTP_DEL_KO",
                    $"Suppression du fichier distant Besançon échouée : {ex.Message}");
            }
        }
    }
}
