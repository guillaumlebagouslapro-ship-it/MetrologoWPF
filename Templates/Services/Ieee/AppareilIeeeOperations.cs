using System;
using System.Collections.Generic;
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
    /// Méthodes d'extension sur <see cref="AppareilIEEE"/> portant la logique Delphi
    /// de <c>TAppareilIEEE.Lecture</c>, <c>ConfigAppareil</c>, et <c>MesureIntervalleRacalDana</c>.
    ///
    /// Toutes les commandes GPIB passent par un <see cref="IIeeeDriver"/> injecté — ce qui
    /// permet de tester en simulation sans matériel et de changer de backend (VISA, NI-488.2, …)
    /// sans toucher à la logique de mesure.
    ///
    /// Référence Delphi : U_DeclarationsMETROLOGO.pas lignes 674 (Lecture), 1047 (MesureIntervalleRacalDana)
    /// et F_Main.pas ligne 1927 (ConfigAppareil).
    /// </summary>
    public static class AppareilIeeeOperations
    {
        // Bit MAV (Message Available) de l'octet de statut IEEE-488.2.
        private const byte BitMav = 0x10;

        /// <summary>
        /// Envoie la chaîne d'initialisation de l'appareil (<c>ChaineInit</c> du .ini).
        /// </summary>
        public static Task InitialiserAsync(
            this AppareilIEEE appareil, IIeeeDriver driver, CancellationToken ct = default)
        {
            if (string.IsNullOrEmpty(appareil.ChaineInit)) return Task.CompletedTask;
            return driver.EcrireAsync(appareil.Adresse, appareil.ChaineInit, appareil.WriteTerm, ct);
        }

        /// <summary>
        /// Portage de <c>TfrmMain.ConfigAppareil</c> : IFC → voie MUX → ConfEntree → activation SRQ.
        /// </summary>
        /// <param name="mux">Multiplexeur optionnel. Ignoré si null ou si <c>mesure.VoieMux &lt;= 0</c>.</param>
        /// <param name="commandesMux">Commandes MUX indexées par numéro de voie (1..N). Ignoré si null.</param>
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
        /// </summary>
        /// <param name="verifierArming">
        /// Si <c>true</c> (défaut), relit l'arming après écriture pour s'assurer que l'instrument
        /// a bien pris en compte les commandes (3 paires query/read = ~200 ms d'overhead). À
        /// désactiver dans les boucles de balayage où chaque ms compte et où l'arming est déjà
        /// validé par le bon comportement des mesures.
        /// </param>
        public static async Task AppliquerGateAsync(
            this AppareilIEEE appareil, IIeeeDriver driver, int gateIndex,
            TypeMesure typeMesure, CancellationToken ct = default,
            bool verifierArming = true)
        {
            if (typeMesure == TypeMesure.Interval) return;

            if (!appareil.Gates.TryGetValue(gateIndex, out var gate))
            {
                // Silence = bug invisible : l'appareil reste en mode gate par défaut et
                // la cadence observée ne correspond pas à ce que l'utilisateur a choisi.
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

            // Le timeout VISA par défaut (5 s) est trop court dès que la gate dépasse ~3 s :
            // le :READ? ne rend la main qu'à la fin du gate, donc il faut que le driver attende
            // au moins gate + marge handshake GPIB. Sinon, premier READ = 0 Hz (chaîne vide après
            // timeout), puis le Write suivant se bloque car le 53131A n'a pas eu le temps de
            // vider son buffer de sortie.
            int timeoutMs = Math.Max(5000, (int)(gate.ValeurSecondes * 1000) + 2000);
            driver.DefinirTimeout(appareil.Adresse, timeoutMs);

            if (string.IsNullOrEmpty(gate.Commande)) return;

            // Certains compteurs (notamment le 53131A) digèrent mal les commandes SCPI chaînées
            // avec ";" dans un seul Write : la 1ère s'applique, les suivantes sont silencieusement
            // ignorées. On scinde donc en writes successifs avec un petit délai pour laisser
            // l'instrument internaliser chaque changement d'état d'arming.
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

            // Vérification : relit les valeurs d'arming pour confirmer que l'instrument a pris
            // en compte nos commandes. Coût : ~200 ms (3 paires query/read). Skippable dans les
            // boucles de balayage où l'arming est déjà éprouvé.
            if (verifierArming)
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

                // On requête aussi l'erreur SCPI en attente — si une commande a été rejetée plus tôt,
                // elle sera dans la file d'erreur (ex : "-113,"Undefined header"").
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
        /// Désactive la génération de SRQ (ex: "QM0" pour Racal, "SR00" pour EIP).
        /// À appeler en fin de boucle de mesures — cf. F_Main.pas:1263 (correctif bug Racal 10ms↔20ms).
        /// </summary>
        public static Task DesactiverSrqAsync(
            this AppareilIEEE appareil, IIeeeDriver driver, CancellationToken ct = default)
        {
            if (!appareil.GereSRQ || string.IsNullOrEmpty(appareil.SRQOff)) return Task.CompletedTask;
            return driver.EcrireAsync(appareil.Adresse, appareil.SRQOff, appareil.WriteTerm, ct);
        }

        /// <summary>
        /// Portage de <c>TAppareilIEEE.Lecture</c> : exécute une mesure unique et retourne la valeur.
        /// Retourne 0.0 en cas de timeout (comportement Delphi identique).
        /// </summary>
        public static async Task<double> MesurerAsync(
            this AppareilIEEE appareil,
            IIeeeDriver driver,
            Mesure mesure,
            CancellationToken ct = default)
        {
            string reponse = await EcrireEtLireAsync(appareil, driver, appareil.ExeMesure, ct);

            if (string.IsNullOrEmpty(reponse)) return 0.0;

            return ParserValeur(reponse, appareil.TailleHeaderReponse);
        }

        /// <summary>
        /// Mode rapide : envoie une commande de fetch (typiquement <c>:FETCh:FREQ?</c>) qui lit
        /// la dernière mesure complétée par l'instrument **sans ré-armer**. À utiliser dans
        /// une boucle avec <c>:INIT:CONT ON</c> activé en amont — sinon retourne toujours la
        /// même valeur. Beaucoup plus rapide que <see cref="MesurerAsync"/> sur le 53131A
        /// (~30-50 ms/mesure vs ~670 ms) car on évite l'arming à chaque appel.
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
        /// Mode bulk : envoie une commande de mesure multiple (ex 53131A :
        /// <c>:SAMP:COUN 30;:READ:ARR? 30</c>) qui demande à l'instrument de faire N
        /// mesures en interne et de les retourner en bloc. Évite N aller-retours GPIB
        /// → gain majeur sur les boucles courtes (gate ≤ 100 ms).
        ///
        /// Le placeholder <c>{N}</c> dans <paramref name="commandeTemplate"/> est remplacé
        /// par <paramref name="nbMesures"/>. La réponse est parsée comme CSV (séparateur
        /// virgule, format SCPI standard — chaque valeur peut avoir son propre header de
        /// taille <c>TailleHeaderReponse</c>, mais on ignore cette subtilité ici car en
        /// pratique les bulk transfers SCPI ne mettent pas de header par valeur).
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

            // Log la réponse brute (tronquée à 200 char pour ne pas saturer le journal) pour
            // pouvoir adapter la commande SCPI si elle ne marche pas avec l'instrument.
            string apercu = string.IsNullOrEmpty(reponse)
                ? "(vide)"
                : reponse.Length > 200 ? reponse.Substring(0, 200) + "..." : reponse;
            JournalLog.Info(CategorieLog.Mesure, "BULK_RAW_RESPONSE",
                $"Cmd envoyée : {commande} | Réponse brute : « {apercu} »");

            var valeurs = ParserValeursMultiples(reponse, nbMesures);
            return valeurs;
        }

        /// <summary>
        /// Parse une réponse SCPI multi-valeurs (CSV ou whitespace-separated). Tolère :
        /// les virgules, points-virgules, espaces et tabs comme séparateurs. Les valeurs
        /// non parsables sont remplacées par 0. Si moins de <paramref name="nbAttendu"/>
        /// valeurs sont reçues, complète avec des 0 (et on logguera côté orchestrator).
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
        /// Dérive une commande de fetch à partir de la commande de mesure de l'appareil.
        /// Convention SCPI : <c>:READ:XXX?</c> ↔ <c>:FETCh:XXX?</c> (READ ré-arme + lit ; FETCH
        /// ne fait que lire la dernière mesure complète). Retourne <c>null</c> si la commande
        /// d'origine ne correspond pas au pattern (ex: appareil pré-SCPI), auquel cas le mode
        /// rapide n'est pas applicable et on retombe sur le READ classique.
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

        /// <summary>
        /// Équivalent Delphi <c>EcritureLectureIEEE</c> : envoi commande + attente MAV (si SRQ géré) + lecture.
        /// </summary>
        private static async Task<string> EcrireEtLireAsync(
            AppareilIEEE appareil, IIeeeDriver driver, string commande, CancellationToken ct)
        {
            await driver.EcrireAsync(appareil.Adresse, commande, appareil.WriteTerm, ct);

            if (appareil.GereSRQ)
            {
                // Poll du MAV jusqu'à disponibilité du résultat.
                while (true)
                {
                    ct.ThrowIfCancellationRequested();
                    var status = await driver.LireStatusByteAsync(appareil.Adresse, ct);
                    if ((status & BitMav) == BitMav) break;
                    await Task.Delay(50, ct);
                }
            }

            return await driver.LireAsync(appareil.Adresse, appareil.ReadTerm, ct);
        }

        /// <summary>
        /// Saute l'entête de la réponse et parse le nombre.
        /// TailleHeaderReponse est la position 1-based du premier caractère numérique
        /// (1 = pas de saut, 3 = saute 2 caractères — ex. entêtes type "F:" sur certains compteurs).
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
