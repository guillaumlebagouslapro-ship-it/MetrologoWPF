using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Metrologo.Models;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Ieee
{
    /// <summary>
    /// Logique de mesure portée du Delphi (TAppareilIEEE.Lecture, ConfigAppareil, MesureIntervalleRacalDana),
    /// via un IIeeeDriver injecté : on peut simuler sans matériel et changer de backend GPIB sans toucher ici.
    /// </summary>
    // réf Delphi : U_DeclarationsMETROLOGO.pas:674 (Lecture), :1047 (MesureIntervalleRacalDana), F_Main.pas:1927 (ConfigAppareil)
    public static class AppareilIeeeOperations
    {
        // bit MAV (Message Available) du status byte IEEE-488.2
        private const byte BitMav = 0x10;

        /// <summary>Envoie la chaîne d'init de l'appareil (ChaineInit du .ini).</summary>
        public static Task InitialiserAsync(
            this AppareilIEEE appareil, IIeeeDriver driver, CancellationToken ct = default)
        {
            if (string.IsNullOrEmpty(appareil.ChaineInit)) return Task.CompletedTask;
            return driver.EcrireAsync(appareil.Adresse, appareil.ChaineInit, appareil.WriteTerm, ct);
        }

        /// <summary>
        /// Portage de TfrmMain.ConfigAppareil : IFC, puis voie MUX (si mux fourni et VoieMux valide),
        /// ConfEntree et activation SRQ.
        /// </summary>
        public static async Task ConfigurerAsync(
            this AppareilIEEE appareil,
            IIeeeDriver driver,
            Mesure mesure,
            AppareilIEEE? mux = null,
            string[]? commandesMux = null,
            CancellationToken ct = default)
        {
            await driver.SendInterfaceClearAsync(ct);

            if (mux != null && commandesMux != null
                && mesure.VoieMux >= 1 && mesure.VoieMux <= commandesMux.Length)
            {
                await driver.EcrireAsync(mux.Adresse, commandesMux[mesure.VoieMux - 1], mux.WriteTerm, ct);
            }

            if (!mesure.InitManu && !string.IsNullOrEmpty(appareil.ConfEntree))
            {
                await driver.EcrireAsync(appareil.Adresse, appareil.ConfEntree, appareil.WriteTerm, ct);
            }

            if (appareil.GereSRQ && !string.IsNullOrEmpty(appareil.SRQOn))
            {
                await driver.EcrireAsync(appareil.Adresse, appareil.SRQOn, appareil.WriteTerm, ct);
            }
        }

        /// <summary>
        /// Programme la gate courante sur l'appareil (sauf en mode Interval, cf. F_Main.pas:1211).
        /// verifierArming relit l'arming après écriture (~200 ms), à couper dans les boucles de balayage.
        /// </summary>
        public static async Task AppliquerGateAsync(
            this AppareilIEEE appareil, IIeeeDriver driver, int gateIndex,
            TypeMesure typeMesure, CancellationToken ct = default,
            bool verifierArming = true)
        {
            if (typeMesure == TypeMesure.Interval) return;

            if (!appareil.Gates.TryGetValue(gateIndex, out var gate))
            {
                // on loggue, sinon bug invisible : l'appareil reste en gate par défaut
                // et la cadence ne correspond pas à ce que l'utilisateur a choisi
                JournalLog.Warn(CategorieLog.Mesure, "GATE_INTROUVABLE",
                    $"Gate index {gateIndex} absent de {appareil.Nom} — commande de gate non envoyée, "
                    + "l'appareil tournera en mode par défaut.",
                    new
                    {
                        Nom = appareil.Nom,
                        GateIndexDemande = gateIndex,
                        GatesDisponibles = appareil.Gates.Keys.OrderBy(k => k).ToArray()
                    });
                return;
            }

            // le timeout VISA par défaut (5 s) est trop court dès que la gate dépasse ~3 s :
            // le :READ? ne rend la main qu'à la fin du gate, il faut donc attendre au moins
            // gate + marge handshake GPIB. Sinon premier READ = 0 Hz (chaîne vide après timeout)
            // puis le Write suivant bloque, le 53131A n'ayant pas vidé son buffer de sortie.
            int timeoutMs = Math.Max(5000, (int)(gate.ValeurSecondes * 1000) + 2000);
            driver.DefinirTimeout(appareil.Adresse, timeoutMs);

            if (string.IsNullOrEmpty(gate.Commande)) return;

            // certains compteurs (le 53131A en particulier) digèrent mal les commandes SCPI
            // chaînées par ";" dans un seul Write : la 1ère passe, les suivantes sont ignorées
            // en silence. On scinde en writes successifs avec un petit délai entre chaque.
            var sousCommandes = gate.Commande
                .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

            foreach (var sc in sousCommandes)
            {
                await driver.EcrireAsync(appareil.Adresse, sc, appareil.WriteTerm, ct);
                JournalLog.Info(CategorieLog.Mesure, "GATE_ENVOI",
                    $"GPIB0::{appareil.Adresse} ← {sc}",
                    new { Nom = appareil.Nom, Adresse = appareil.Adresse, Commande = sc });
                await Task.Delay(50, ct);
            }

            // relecture de l'arming pour confirmer la prise en compte (~200 ms, 3 paires query/read).
            // Réservé aux modèles qui supportent :FREQ:ARM:* (HP/Agilent 53131A et compatibles) :
            // les autres (53230A, SR620...) renvoient -113 Undefined header, qui s'affiche à
            // l'écran de l'instrument. D'où le champ catalogue VerifArmingActive (défaut false).
            if (verifierArming && appareil.VerifArmingActive)
            {
                await VerifierArmingAsync(appareil, driver, ct);
            }
        }

        private static async Task VerifierArmingAsync(
            AppareilIEEE appareil, IIeeeDriver driver, CancellationToken ct)
        {
            try
            {
                await driver.EcrireAsync(appareil.Adresse, ":FREQ:ARM:STOP:SOUR?", appareil.WriteTerm, ct);
                var sour = (await driver.LireAsync(appareil.Adresse, appareil.ReadTerm, ct))?.Trim() ?? "";

                await driver.EcrireAsync(appareil.Adresse, ":FREQ:ARM:STOP:TIM?", appareil.WriteTerm, ct);
                var tim = (await driver.LireAsync(appareil.Adresse, appareil.ReadTerm, ct))?.Trim() ?? "";

                // on lit aussi la file d'erreur SCPI : une commande rejetée plus tôt
                // y traîne encore (ex -113 Undefined header)
                await driver.EcrireAsync(appareil.Adresse, ":SYST:ERR?", appareil.WriteTerm, ct);
                var err = (await driver.LireAsync(appareil.Adresse, appareil.ReadTerm, ct))?.Trim() ?? "";

                JournalLog.Info(CategorieLog.Mesure, "GATE_VERIF",
                    $"Arming relu : STOP:SOUR={sour} · STOP:TIM={tim} · SYST:ERR={err}",
                    new { StopSour = sour, StopTim = tim, SystErr = err });
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Mesure, "GATE_VERIF_ECHEC",
                    $"Impossible de relire l'arming sur {appareil.Nom} : {ex.GetType().Name} — {ex.Message}");
            }
        }

        /// <summary>
        /// Coupe la génération de SRQ ("QM0" sur Racal, "SR00" sur EIP) en fin de boucle de mesures.
        /// Cf. F_Main.pas:1263 (correctif du bug Racal 10 ms / 20 ms).
        /// </summary>
        public static Task DesactiverSrqAsync(
            this AppareilIEEE appareil, IIeeeDriver driver, CancellationToken ct = default)
        {
            if (!appareil.GereSRQ || string.IsNullOrEmpty(appareil.SRQOff)) return Task.CompletedTask;
            return driver.EcrireAsync(appareil.Adresse, appareil.SRQOff, appareil.WriteTerm, ct);
        }

        /// <summary>
        /// Portage de TAppareilIEEE.Lecture : une mesure unique. Retourne 0.0 sur timeout, comme en Delphi.
        /// </summary>
        public static async Task<double> MesurerAsync(
            this AppareilIEEE appareil,
            IIeeeDriver driver,
            Mesure mesure,
            CancellationToken ct = default,
            int mavTimeoutMs = 10000)
        {
            string reponse = await EcrireEtLireAsync(appareil, driver, appareil.ExeMesure, ct, mavTimeoutMs);

            if (string.IsNullOrEmpty(reponse)) return 0.0;

            return ParserValeur(reponse, appareil.TailleHeaderReponse);
        }

        /// <summary>
        /// Mode rapide : commande de fetch (typiquement :FETCh:FREQ?) qui lit la dernière mesure
        /// sans ré-armer. Suppose :INIT:CONT ON activé en amont, sinon on relit toujours la même
        /// valeur. Sur 53131A : ~30-50 ms/mesure contre ~670 ms pour un READ complet.
        /// </summary>
        public static async Task<double> FetcherAsync(
            this AppareilIEEE appareil,
            IIeeeDriver driver,
            string commandeFetch,
            CancellationToken ct = default)
        {
            if (string.IsNullOrWhiteSpace(commandeFetch)) return 0.0;
            string reponse = await EcrireEtLireAsync(appareil, driver, commandeFetch, ct);
            if (string.IsNullOrEmpty(reponse)) return 0.0;
            return ParserValeur(reponse, appareil.TailleHeaderReponse);
        }

        /// <summary>
        /// Mode bulk : l'instrument fait N mesures en interne et les renvoie en bloc
        /// (ex 53131A : ":SAMP:COUN 30;:READ:ARR? 30"), ce qui évite N aller-retours GPIB.
        /// Gros gain sur les gates courtes (100 ms et moins). Le placeholder {N} du template
        /// est remplacé par nbMesures, la réponse est parsée en CSV sans header par valeur.
        /// </summary>
        public static async Task<List<double>> MesurerEnLotAsync(
            this AppareilIEEE appareil,
            IIeeeDriver driver,
            string commandeTemplate,
            int nbMesures,
            CancellationToken ct = default)
        {
            string commande = commandeTemplate.Replace("{N}",
                nbMesures.ToString(CultureInfo.InvariantCulture));

            string reponse = await EcrireEtLireAsync(appareil, driver, commande, ct);

            // on loggue la réponse brute (tronquée à 200 car.) pour pouvoir adapter
            // la commande SCPI si l'instrument ne la digère pas
            string apercu = string.IsNullOrEmpty(reponse)
                ? "(vide)"
                : reponse.Length > 200 ? reponse.Substring(0, 200) + "..." : reponse;
            JournalLog.Info(CategorieLog.Mesure, "BULK_RAW_RESPONSE",
                $"Cmd envoyée : {commande} | Réponse brute : « {apercu} »");

            var valeurs = ParserValeursMultiples(reponse, nbMesures);
            return valeurs;
        }

        /// <summary>
        /// Parse une réponse SCPI multi-valeurs. Séparateurs tolérés : virgule, point-virgule,
        /// espace, tab. Les morceaux non parsables sont ignorés, à l'appelant de vérifier le compte.
        /// </summary>
        public static List<double> ParserValeursMultiples(string reponse, int nbAttendu)
        {
            var resultat = new List<double>(nbAttendu);
            if (string.IsNullOrWhiteSpace(reponse)) return resultat;

            var separateurs = new[] { ',', ';', ' ', '\t', '\r', '\n' };
            var parts = reponse.Split(separateurs, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                if (double.TryParse(part.Trim(), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var v))
                {
                    resultat.Add(v);
                }
            }
            return resultat;
        }

        /// <summary>
        /// Dérive la commande de fetch depuis la commande de mesure : :READ:XXX? devient :FETCh:XXX?
        /// (READ ré-arme + lit, FETCH lit juste la dernière mesure). Retourne null si pas de pattern
        /// READ (appareil pré-SCPI) : pas de mode rapide possible, on garde le READ classique.
        /// </summary>
        public static string? DeriverCommandeFetch(string commandeMesure)
        {
            if (string.IsNullOrWhiteSpace(commandeMesure)) return null;
            string trim = commandeMesure.Trim();

            // Cas usuels SCPI :
            //   :READ:FREQ?    -> :FETCh:FREQ?
            //   :READ?         -> :FETCh?
            //   READ:FREQ?     -> FETCh:FREQ?
            if (trim.StartsWith(":READ", System.StringComparison.OrdinalIgnoreCase))
                return ":FETCh" + trim.Substring(":READ".Length);
            if (trim.StartsWith("READ", System.StringComparison.OrdinalIgnoreCase))
                return "FETCh" + trim.Substring("READ".Length);

            return null;
        }

        // ---------------- Interne ----------------

        /// <summary>Équivalent Delphi EcritureLectureIEEE : envoi commande + attente MAV (si SRQ géré) + lecture.</summary>
        private static async Task<string> EcrireEtLireAsync(
            AppareilIEEE appareil, IIeeeDriver driver, string commande, CancellationToken ct,
            int mavTimeoutMs = 10000)
        {
            JournalLog.Info(CategorieLog.Mesure, "MESURE_ENVOI",
                $"GPIB0::{appareil.Adresse} ← {commande}",
                new { Nom = appareil.Nom, Adresse = appareil.Adresse, Commande = commande });

            await driver.EcrireAsync(appareil.Adresse, commande, appareil.WriteTerm, ct);

            if (appareil.GereSRQ)
            {
                // Attente du MAV (Message Available) AVANT bornage : certains compteurs legacy
                // (EIP 545…) ne positionnent pas toujours le bit MAV via le serial-poll du driver
                // VISA (comportement différent du NI488 d'origine). Un poll infini bloquait alors
                // la mesure sur le 1er point (Excel ouvert, aucune valeur). On BORNE donc l'attente
                // (timeout calé sur la gate) puis on lit quand même : si la donnée est déjà dans le
                // buffer de sortie, la lecture la récupère ; sinon timeout VISA → chaîne vide → 0.0,
                // sans jamais figer la mesure.
                var sw = Stopwatch.StartNew();
                bool mav = false;
                while (sw.ElapsedMilliseconds < mavTimeoutMs)
                {
                    ct.ThrowIfCancellationRequested();
                    byte status;
                    try { status = await driver.LireStatusByteAsync(appareil.Adresse, ct); }
                    catch (OperationCanceledException) { throw; }
                    catch { break; }   // serial-poll non supporté → on tente la lecture directe
                    if ((status & BitMav) == BitMav) { mav = true; break; }
                    await Task.Delay(50, ct);
                }
                if (!mav)
                    JournalLog.Warn(CategorieLog.Mesure, "MAV_TIMEOUT",
                        $"GPIB0::{appareil.Adresse} : MAV non positionné après {mavTimeoutMs} ms "
                      + $"(cmd « {commande} ») — lecture directe tentée.");
            }

            string reponse = await driver.LireAsync(appareil.Adresse, appareil.ReadTerm, ct);

            // Trace la réponse brute (essentielle pour diagnostiquer un legacy muet) : on voit
            // immédiatement si l'appareil renvoie une valeur, une chaîne vide ou un format inattendu.
            string apercu = string.IsNullOrEmpty(reponse) ? "(vide)"
                : reponse.Length > 120 ? reponse.Substring(0, 120) + "..." : reponse;
            JournalLog.Info(CategorieLog.Mesure, "MESURE_RECEP",
                $"GPIB0::{appareil.Adresse} → « {apercu} »");

            return reponse;
        }

        /// <summary>
        /// Saute l'entête de la réponse et parse le nombre. TailleHeaderReponse est la position
        /// 1-based du premier caractère numérique (1 = pas de saut, 3 = saute 2 caractères,
        /// genre entête "F:" sur certains compteurs).
        /// </summary>
        private static double ParserValeur(string reponse, int tailleHeader)
        {
            int saut = Math.Max(0, tailleHeader - 1);
            if (saut >= reponse.Length) return 0.0;

            string nombre = reponse[saut..].Trim();
            return double.TryParse(nombre, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)
                ? v
                : 0.0;
        }
    }
}
