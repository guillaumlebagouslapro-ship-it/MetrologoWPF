using System.Diagnostics;
using System.Globalization;
using Metrologo.Models;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Ieee
{
    /// <summary>
    /// Logique de mesure reprise du Delphi (TAppareilIEEE.Lecture, ConfigAppareil,
    /// MesureIntervalleRacalDana), via un IIeeeDriver injecté (simulation ou backend GPIB au choix).
    /// </summary>
    // réf Delphi : U_DeclarationsMETROLOGO.pas:674 (Lecture), :1047 (MesureIntervalleRacalDana), F_Main.pas:1927 (ConfigAppareil)
    public static class AppareilIeeeOperations
    {

        /// <summary>Envoie la chaîne d'initialisation de l'appareil (le ChaineInit du .ini).</summary>
        public static Task InitialiserAsync(
            this AppareilIEEE appareil, IIeeeDriver driver, CancellationToken ct = default)
        {
            if (string.IsNullOrEmpty(appareil.ChaineInit)) return Task.CompletedTask;
            return driver.EcrireAsync(appareil.Adresse, appareil.ChaineInit, appareil.WriteTerm, ct);
        }

        /// <summary>
        /// Reprise de TfrmMain.ConfigAppareil : IFC, voie MUX (si mux fourni et VoieMux valide),
        /// ConfEntree, puis activation du SRQ.
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
        /// Programme la gate en cours sur l'appareil, sauf en mode Interval (cf. F_Main.pas:1211).
        /// verifierArming relit l'arming après l'écriture (~200 ms) : à désactiver dans les boucles de balayage.
        /// </summary>
        public static async Task AppliquerGateAsync(
            this AppareilIEEE appareil, IIeeeDriver driver, int gateIndex,
            TypeMesure typeMesure, CancellationToken ct = default,
            bool verifierArming = true)
        {
            if (typeMesure == TypeMesure.Interval) return;

            if (!appareil.Gates.TryGetValue(gateIndex, out var gate))
            {
                // Tracé, sinon invisible : l'appareil reste sur sa gate par défaut.
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

            // Le :READ? ne répond qu'une fois la gate terminée : timeout >= gate + marge (5 s par
            // défaut, trop juste au-delà de ~3 s). Sinon le READ rend 0 Hz (vide après timeout) et le
            // Write suivant se bloque, le 53131A n'ayant pas vidé son buffer de sortie.
            int timeoutMs = Math.Max(5000, (int)(gate.ValeurSecondes * 1000) + 2000);
            driver.DefinirTimeout(appareil.Adresse, timeoutMs);

            if (string.IsNullOrEmpty(gate.Commande)) return;

            // Le 53131A (entre autres) ignore sans rien dire les commandes après le premier ";" d'un
            // même Write : on envoie donc chaque sous-commande séparément, avec un petit délai.
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

            // Relecture de l'arming (~200 ms), réservée aux modèles qui connaissent :FREQ:ARM:*
            // (53131A et compatibles) : les autres (53230A, SR620...) répondent -113 Undefined header,
            // affiché sur l'instrument. D'où le champ catalogue VerifArmingActive (false par défaut).
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

                // Vide aussi la file d'erreurs SCPI (ex. un -113 laissé par une commande rejetée).
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
        /// Coupe la génération du SRQ en fin de boucle de mesures ("QM0" sur Racal, "SR00" sur EIP).
        /// Cf. F_Main.pas:1263, le correctif du bug Racal 10 ms / 20 ms.
        /// </summary>
        public static Task DesactiverSrqAsync(
            this AppareilIEEE appareil, IIeeeDriver driver, CancellationToken ct = default)
        {
            if (!appareil.GereSRQ || string.IsNullOrEmpty(appareil.SRQOff)) return Task.CompletedTask;
            return driver.EcrireAsync(appareil.Adresse, appareil.SRQOff, appareil.WriteTerm, ct);
        }

        /// <summary>
        /// Reprise de TAppareilIEEE.Lecture : une mesure. Comme en Delphi, renvoie 0.0 en cas de timeout.
        /// </summary>
        public static async Task<double> MesurerAsync(
            this AppareilIEEE appareil,
            IIeeeDriver driver,
            CancellationToken ct = default,
            int gateMs = 0)
        {
            string reponse = await EcrireEtLireAsync(appareil, driver, appareil.ExeMesure, ct, gateMs);

            if (string.IsNullOrEmpty(reponse)) return 0.0;

            return ParserValeur(reponse, appareil.TailleHeaderReponse);
        }

        /// <summary>
        /// Mode rapide : fetch (typiquement :FETCh:FREQ?) de la dernière mesure sans ré-armer.
        /// Suppose :INIT:CONT ON en amont, sinon la même valeur est relue. 53131A : ~30-50 ms
        /// par mesure contre ~670 ms pour un READ complet.
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
        /// (ex 53131A : ":SAMP:COUN 30;:READ:ARR? 30"), d'où un gros gain sur les gates courtes.
        /// {N} est remplacé par nbMesures ; la réponse est un CSV sans header par valeur.
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

            // Réponse brute tronquée journalisée, pour adapter la commande si l'instrument la refuse.
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
        /// Dérive la commande de fetch de la commande de mesure : :READ:XXX? devient :FETCh:XXX?
        /// (READ ré-arme puis lit, FETCH lit juste la dernière mesure). Null si pas de pattern READ
        /// (appareil pré-SCPI) : pas de mode rapide.
        /// </summary>
        public static string? DeriverCommandeFetch(string commandeMesure)
        {
            if (string.IsNullOrWhiteSpace(commandeMesure)) return null;
            string trim = commandeMesure.Trim();

            // :READ:FREQ? -> :FETCh:FREQ? ; :READ? -> :FETCh? ; READ:FREQ? -> FETCh:FREQ?
            if (trim.StartsWith(":READ", System.StringComparison.OrdinalIgnoreCase))
                return ":FETCh" + trim.Substring(":READ".Length);
            if (trim.StartsWith("READ", System.StringComparison.OrdinalIgnoreCase))
                return "FETCh" + trim.Substring("READ".Length);

            return null;
        }

        // ---------------- Interne ----------------

        /// <summary>Équivalent Delphi EcritureLectureIEEE : envoi, attente « mesure prête » (appareils
        /// à SRQ) puis lecture. gateMs = durée de gate en ms (0 = inconnue), sert au garde-fou temps.</summary>
        private static async Task<string> EcrireEtLireAsync(
            AppareilIEEE appareil, IIeeeDriver driver, string commande, CancellationToken ct,
            int gateMs = 0)
        {
            JournalLog.Trace(CategorieLog.Mesure, "MESURE_ENVOI",
                $"GPIB0::{appareil.Adresse} ← {commande}",
                new { Nom = appareil.Nom, Adresse = appareil.Adresse, Commande = commande });

            await driver.EcrireAsync(appareil.Adresse, commande, appareil.WriteTerm, ct);

            // Handshake SRQ fidèle au Delphi (EcritureLectureIEEE, bGereSRQ=true) : « mesure prête »
            // se détecte par serial poll (ReadStatusByte), pas par un délai aveugle. Le poll lit
            // l'octet de statut sans toucher à la donnée, donc ne perturbe pas la mesure de l'EIP.
            // Prêt dès que RQS (0x40, appareils pré-488.2 comme l'EIP 545) ou MAV (0x10, 488.2).
            //
            // La mesure dure au moins la gate : on dort d'abord la gate en une fois (pas de trafic
            // bus, un poll coûte ~30-50 ms en USB-GPIB) puis on confirme par poll. Borné à gate +
            // marge : si le SRQ n'arrive jamais, on lit quand même plutôt que de bloquer.
            if (appareil.GereSRQ)
            {
                const byte PRET = 0x40 | 0x10;   // RQS | MAV
                if (gateMs > 0) await Task.Delay(gateMs, ct);

                int budgetMs = gateMs > 0 ? gateMs + 2000 : 4000;
                var swSrq = Stopwatch.StartNew();
                byte statut = 0;
                bool pret = false;
                while (swSrq.ElapsedMilliseconds < budgetMs)
                {
                    ct.ThrowIfCancellationRequested();
                    statut = await driver.LireStatusByteAsync(appareil.Adresse, ct);
                    if ((statut & PRET) != 0) { pret = true; break; }
                    await Task.Delay(5, ct);
                }
                if (!pret)
                    JournalLog.Warn(CategorieLog.Mesure, "SRQ_TIMEOUT",
                        $"GPIB0::{appareil.Adresse} statut=0x{statut:X2} après {swSrq.ElapsedMilliseconds} ms "
                        + $"(budget {budgetMs} ms) — lecture quand même.");
                else
                    JournalLog.Trace(CategorieLog.Mesure, "SRQ_OK",
                        $"GPIB0::{appareil.Adresse} statut=0x{statut:X2} confirmé en {swSrq.ElapsedMilliseconds} ms.");
            }

            string reponse = await driver.LireAsync(appareil.Adresse, appareil.ReadTerm, ct);

            // Réponse brute en trace (niveau verbeux) : diagnostic d'un appareil legacy muet.
            string apercu = string.IsNullOrEmpty(reponse) ? "(vide)"
                : reponse.Length > 120 ? reponse.Substring(0, 120) + "..." : reponse;
            JournalLog.Trace(CategorieLog.Mesure, "MESURE_RECEP",
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
