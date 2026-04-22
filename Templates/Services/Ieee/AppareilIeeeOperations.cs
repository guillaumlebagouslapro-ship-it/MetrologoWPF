using System;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Metrologo.Models;

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
        // Nom canonique du Racal — permet de détecter le cas spécial mesure d'intervalle.
        public const string NomRacal = "Racal-Dana 1996";

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
        public static Task AppliquerGateAsync(
            this AppareilIEEE appareil, IIeeeDriver driver, int gateIndex,
            TypeMesure typeMesure, CancellationToken ct = default)
        {
            if (typeMesure == TypeMesure.Interval) return Task.CompletedTask;
            if (!appareil.Gates.TryGetValue(gateIndex, out var gate)) return Task.CompletedTask;
            if (string.IsNullOrEmpty(gate.Commande)) return Task.CompletedTask;

            return driver.EcrireAsync(appareil.Adresse, gate.Commande, appareil.WriteTerm, ct);
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
            string reponse;

            if (appareil.Nom == NomRacal && mesure.TypeMesure == TypeMesure.Interval)
            {
                reponse = await MesurerIntervalleRacalAsync(appareil, driver, ct);
            }
            else
            {
                reponse = await EcrireEtLireAsync(appareil, driver, appareil.ExeMesure, ct);
            }

            if (string.IsNullOrEmpty(reponse)) return 0.0;

            return ParserValeur(reponse, appareil.TailleHeaderReponse);
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
        /// Portage fidèle de <c>MesureIntervalleRacalDana</c> (U_DeclarationsMETROLOGO.pas:1047).
        /// </summary>
        private static async Task<string> MesurerIntervalleRacalAsync(
            AppareilIEEE appareil, IIeeeDriver driver, CancellationToken ct)
        {
            // 1. Inhibe SRQ et draine les messages en attente.
            await driver.EcrireAsync(appareil.Adresse, "QM0", appareil.WriteTerm, ct);

            while (true)
            {
                ct.ThrowIfCancellationRequested();
                var status = await driver.LireStatusByteAsync(appareil.Adresse, ct);
                if ((status & BitMav) != BitMav) break;
                await driver.LireAsync(appareil.Adresse, appareil.ReadTerm, ct);  // drain
            }

            // 2. Déclenche une nouvelle mesure en mode local.
            await driver.EcrireAsync(appareil.Adresse, "RE", appareil.WriteTerm, ct);
            await driver.DefinirRemoteLocalAsync(appareil.Adresse, remote: false, ct);

            // 3. Attend la disponibilité du résultat (polling toutes les 1s comme Delphi).
            while (true)
            {
                ct.ThrowIfCancellationRequested();
                var status = await driver.LireStatusByteAsync(appareil.Adresse, ct);
                if ((status & BitMav) == BitMav) break;
                await Task.Delay(1000, ct);
            }

            return await driver.LireAsync(appareil.Adresse, appareil.ReadTerm, ct);
        }

        /// <summary>
        /// Saute l'entête de la réponse et parse le nombre.
        /// TailleHeaderReponse est la position 1-based du premier caractère numérique
        /// (Stanford=1 → pas de saut ; Racal=3 → saute 2 caractères).
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
