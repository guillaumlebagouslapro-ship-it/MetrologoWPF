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
        /// <summary>Résultat d'un téléchargement FTP : contenu si OK, sinon détail de l'échec.</summary>
        public sealed class ResultatFtp
        {
            public string? Contenu { get; set; }
            public bool ConfigManquante { get; set; }
            public string Url { get; set; } = "";
            public string? Erreur { get; set; }
            public bool Ok => Contenu != null;
        }

        /// <summary>
        /// Télécharge le fichier distant. Retourne le contenu si OK, sinon un <see cref="ResultatFtp"/>
        /// décrivant précisément l'échec : config manquante, ou réponse FTP du serveur (ex.
        /// « 530 Login incorrect », « 550 File not found »), ou erreur réseau (timeout, TLS, refus).
        /// </summary>
        public static async Task<ResultatFtp> TelechargerAsync(BesanconConfig cfg)
        {
            if (string.IsNullOrWhiteSpace(cfg.FtpHote) || string.IsNullOrWhiteSpace(cfg.FtpUtilisateur))
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_FTP_CONFIG",
                    "FTP Besançon non configuré (hôte ou identifiant manquant) — téléchargement ignoré. "
                  + $"Renseigne {BesanconConfig.Chemin}.");
                return new ResultatFtp
                {
                    ConfigManquante = true,
                    Erreur = $"hôte ou identifiant manquant — renseigne {BesanconConfig.Chemin}",
                };
            }

            string url = $"ftp://{cfg.FtpHote}:{cfg.FtpPort}/{cfg.FichierDistant}";
            try
            {
                string contenu = await Task.Run(() =>
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
                return new ResultatFtp { Contenu = contenu, Url = url };
            }
            catch (WebException wex)
            {
                // Réponse FTP du serveur (code + texte) si disponible — le plus parlant pour diagnostiquer.
                string detail = wex.Message;
                if (wex.Response is FtpWebResponse fr)
                {
                    string desc = (fr.StatusDescription ?? "").Trim();
                    detail = $"{(int)fr.StatusCode} {desc} ({wex.Status})".Trim();
                }
                else
                {
                    detail = $"{wex.Message} ({wex.Status})";
                }
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_FTP_KO",
                    $"Téléchargement FTP Besançon échoué ({url}) : {detail}");
                return new ResultatFtp { Url = url, Erreur = detail };
            }
            catch (Exception ex)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_FTP_KO",
                    $"Téléchargement FTP Besançon échoué ({url}) : {ex.Message}");
                return new ResultatFtp { Url = url, Erreur = ex.Message };
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
