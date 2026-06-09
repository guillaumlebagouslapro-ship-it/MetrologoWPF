using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Metrologo.Models;
using Metrologo.Services.Journal;

namespace Metrologo.Services.Besancon
{
    /// <summary>
    /// ⚠ SONDE TEMPORAIRE DE DIAGNOSTIC — à RETIRER une fois l'heure de publication connue.
    ///
    /// <para/>On ne sait pas à quelle heure l'observatoire dépose une nouvelle valeur sur le FTP.
    /// Cette sonde télécharge donc le fichier <b>chaque minute</b>, repère l'apparition d'un nouveau
    /// MJD (date la plus récente) et <b>journalise l'heure exacte</b> de cette apparition — sans
    /// rien intégrer (pas d'écriture dans <see cref="BesanconTxtStore"/>, pas d'injection de
    /// référence rubidium). Elle COMPLÈTE la tâche quotidienne <see cref="BesanconScheduler"/>
    /// (09h38) qui, elle, fait le vrai traitement.
    ///
    /// <para/>Une fois l'heure de publication observée (voir le fichier
    /// <c>sonde_besancon_detections.csv</c> sur le partage Besançon), il suffira de régler
    /// <see cref="BesanconConfig.HeureDeclenchement"/> sur cette heure puis de supprimer ce
    /// fichier et son appel dans <c>App.xaml.cs</c>.
    /// </summary>
    public static class BesanconSonde
    {
        /// <summary>Intervalle de sondage (legacy : « toutes les minutes »).</summary>
        private static readonly TimeSpan Intervalle = TimeSpan.FromMinutes(1);

        private static Timer? _timer;
        private static readonly object _sync = new();

        /// <summary>Garde-fou anti-chevauchement : un sondage en cours empêche le suivant de démarrer
        /// (un téléchargement FTP peut durer jusqu'à 30 s, soit moins que l'intervalle, mais on
        /// se protège des pics de latence réseau).</summary>
        private static int _enCours;

        /// <summary>Dernière erreur FTP journalisée — pour ne pas spammer le Journal chaque minute :
        /// on ne reloggue qu'au CHANGEMENT d'erreur (ou au retour à la normale).</summary>
        private static string? _derniereErreur;

        /// <summary>Démarre la sonde (no-op si le FTP n'est pas configuré sur ce poste).</summary>
        public static void Demarrer()
        {
            var cfg = BesanconConfig.Charger();
            if (string.IsNullOrWhiteSpace(cfg.FtpHote) || string.IsNullOrWhiteSpace(cfg.FtpUtilisateur))
            {
                Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_SONDE_INACTIVE",
                    "Sonde Besançon non démarrée : FTP non configuré.");
                return;
            }

            lock (_sync)
            {
                _timer?.Dispose();
                // 1er sondage immédiat (établit la ligne de base au lancement), puis chaque minute.
                _timer = new Timer(_ => _ = SonderAsync(), null, TimeSpan.Zero, Intervalle);
            }
            Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_SONDE_DEMARREE",
                $"Sonde Besançon démarrée : sondage FTP toutes les {Intervalle.TotalMinutes:0} min "
              + $"pour repérer l'heure d'apparition d'une nouvelle valeur. Détections journalisées dans {CheminDetections}.");
        }

        /// <summary>Arrête la sonde (dispose le timer s'il tourne).</summary>
        public static void Arreter()
        {
            lock (_sync)
            {
                _timer?.Dispose();
                _timer = null;
            }
        }

        private static async Task SonderAsync()
        {
            // Un seul sondage à la fois.
            if (Interlocked.Exchange(ref _enCours, 1) == 1) return;
            try
            {
                var cfg = BesanconConfig.Charger();
                var ftp = await BesanconFtpService.TelechargerAsync(cfg);
                if (!ftp.Ok)
                {
                    // Ne journalise qu'au changement d'erreur (évite 1 ligne/minute).
                    if (_derniereErreur != ftp.Erreur)
                    {
                        _derniereErreur = ftp.Erreur;
                        Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_SONDE_FTP_KO",
                            $"Sonde Besançon : téléchargement FTP échoué — {ftp.Erreur}");
                    }
                    return;
                }
                // Retour à la normale après une série d'échecs.
                if (_derniereErreur != null)
                {
                    _derniereErreur = null;
                    Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_SONDE_FTP_OK",
                        "Sonde Besançon : FTP de nouveau joignable.");
                }

                var mesures = BesanconParser.Parser(ftp.Contenu!);
                if (mesures.Count == 0) return;

                int maxMjd = mesures.Max(m => m.Mjd);
                double valeur = mesures.First(m => m.Mjd == maxMjd).Valeur;

                var etat = ChargerEtat();
                if (etat == null)
                {
                    // Première observation : on fixe la ligne de base sans crier « nouvelle valeur »
                    // (on ne connaît pas l'heure d'apparition de ce MJD-là, déjà présent au lancement).
                    SauvegarderEtat(maxMjd);
                    Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_SONDE_INIT",
                        $"Sonde Besançon : ligne de base établie au MJD {maxMjd} (déjà présent). "
                      + "Surveillance de la prochaine nouvelle valeur en cours.");
                    return;
                }

                if (maxMjd > etat.DernierMjdVu)
                {
                    var maintenant = DateTime.Now;
                    EnregistrerDetection(maintenant, maxMjd, etat.DernierMjdVu, valeur);
                    SauvegarderEtat(maxMjd);
                    Journal.Journal.Info(CategorieLog.Systeme, "BESANCON_SONDE_NOUVELLE",
                        $"Sonde Besançon : NOUVELLE valeur disponible à {maintenant:dd/MM/yyyy HH:mm:ss} "
                      + $"— MJD {maxMjd} (précédent {etat.DernierMjdVu}), valeur {valeur.ToString(System.Globalization.CultureInfo.InvariantCulture)} Hz.");
                    AfficherPopup(maintenant, maxMjd, etat.DernierMjdVu, valeur);
                }
            }
            catch (Exception ex)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_SONDE_KO",
                    $"Sonde Besançon : erreur de sondage — {ex.Message}");
            }
            finally
            {
                Interlocked.Exchange(ref _enCours, 0);
            }
        }

        /// <summary>
        /// Affiche un pop-up modal BIEN VISIBLE (qu'il faut acquitter) à la détection d'une nouvelle
        /// valeur, en plus du Journal et du CSV. Marshalé sur le thread UI en <c>BeginInvoke</c> :
        /// la boîte est modale côté UI mais ne bloque PAS le thread de la sonde (le sondage suivant
        /// peut repartir). No-op si aucune UI n'est présente (exécution sans fenêtre).
        /// </summary>
        private static void AfficherPopup(DateTime quand, int nouveauMjd, int ancienMjd, double valeur)
        {
            var disp = System.Windows.Application.Current?.Dispatcher;
            if (disp == null) return;

            string message =
                "Une NOUVELLE valeur Besançon vient d'apparaître sur le FTP.\n\n"
              + $"🕒 Heure d'apparition : {quand:HH:mm:ss}  ({quand:dd/MM/yyyy})\n"
              + $"📅 MJD : {nouveauMjd}  (précédent : {ancienMjd})\n"
              + $"📈 Valeur : {valeur.ToString(System.Globalization.CultureInfo.InvariantCulture)} Hz\n\n"
              + $"Heure notée dans :\n{CheminDetections}";

            disp.BeginInvoke(new Action(() =>
            {
                try
                {
                    System.Windows.MessageBox.Show(
                        message,
                        "⚠  NOUVELLE VALEUR BESANÇON DISPONIBLE",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                }
                catch { /* best-effort : ne jamais faire tomber la sonde pour un souci d'affichage */ }
            }));
        }

        // ─── État persistant (survit à un redémarrage : pas de fausse « nouvelle valeur » au lancement) ───

        private static string CheminEtat =>
            Path.Combine(CheminsMetrologo.Besancon, "sonde_besancon_etat.json");

        /// <summary>Journal des détections, en CSV (point-virgule, lisible dans Excel FR).</summary>
        private static string CheminDetections =>
            Path.Combine(CheminsMetrologo.Besancon, "sonde_besancon_detections.csv");

        private sealed class EtatSonde
        {
            public int DernierMjdVu { get; set; }
        }

        private static EtatSonde? ChargerEtat()
        {
            try
            {
                if (File.Exists(CheminEtat))
                    return JsonSerializer.Deserialize<EtatSonde>(File.ReadAllText(CheminEtat));
            }
            catch { /* corrompu ou injoignable → traité comme première observation */ }
            return null;
        }

        private static void SauvegarderEtat(int maxMjd)
        {
            try
            {
                Directory.CreateDirectory(CheminsMetrologo.Besancon);
                File.WriteAllText(CheminEtat,
                    JsonSerializer.Serialize(new EtatSonde { DernierMjdVu = maxMjd },
                        new JsonSerializerOptions { WriteIndented = true }));
            }
            catch (Exception ex)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_SONDE_ETAT_KO",
                    $"Sonde Besançon : écriture de l'état échouée ({CheminEtat}) — {ex.Message}");
            }
        }

        /// <summary>Ajoute une ligne au CSV de diagnostic (crée l'en-tête si le fichier n'existe pas).</summary>
        private static void EnregistrerDetection(DateTime quand, int nouveauMjd, int ancienMjd, double valeur)
        {
            try
            {
                Directory.CreateDirectory(CheminsMetrologo.Besancon);
                bool nouveau = !File.Exists(CheminDetections);
                var sb = new StringBuilder();
                if (nouveau)
                    sb.AppendLine("DateDetection;Heure;MJD;MJDPrecedent;Poste;Valeur(Hz)");
                sb.Append(quand.ToString("dd/MM/yyyy")).Append(';')
                  .Append(quand.ToString("HH:mm:ss")).Append(';')
                  .Append(nouveauMjd).Append(';')
                  .Append(ancienMjd).Append(';')
                  .Append(Environment.MachineName).Append(';')
                  .Append(valeur.ToString(System.Globalization.CultureInfo.InvariantCulture))
                  .Append('\n');
                File.AppendAllText(CheminDetections, sb.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_SONDE_CSV_KO",
                    $"Sonde Besançon : écriture du CSV de détection échouée ({CheminDetections}) — {ex.Message}");
            }
        }
    }
}
