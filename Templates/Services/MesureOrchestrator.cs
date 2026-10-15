using System.Diagnostics;
using System.Globalization;
using System.IO;
using Metrologo.Models;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Ieee;
using Metrologo.Services.Incertitude;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services
{
    public class ResultatMesure
    {
        public bool Succes { get; set; }
        public List<double> Valeurs { get; set; } = new();
        public double? FNominale { get; set; }
        public string? Erreur { get; set; }
        public string? CheminExcel { get; set; }

        /// <summary>Transfert du dossier FI vers M:\ en fin de mesure : null si non tenté, false en cas d'échec
        /// (la FI est alors mise en liste de reprise pour le prochain démarrage).</summary>
        public bool? TransfertReseauOk { get; set; }

        /// <summary>Faux si le rapport devait sortir en .xls (réglage Admin) et que la conversion a échoué
        /// (il reste en .xlsx). Null si aucune conversion n'était demandée.</summary>
        public bool? ConversionXlsOk { get; set; }

        public double Moyenne => Valeurs.Count == 0 ? 0 : Moy(Valeurs);
        public double EcartType => Valeurs.Count < 2 ? 0 : Sigma(Valeurs);

        private static double Moy(List<double> v)
        {
            double s = 0; foreach (var x in v) s += x; return s / v.Count;
        }

        private static double Sigma(List<double> v)
        {
            var m = Moy(v);
            double s = 0; foreach (var x in v) s += (x - m) * (x - m);
            return Math.Sqrt(s / (v.Count - 1));
        }
    }

    public class ProgressionMesure
    {
        public string Message { get; set; } = string.Empty;
        public int EtapeActuelle { get; set; }
        public int EtapesTotales { get; set; }
        public double? DerniereValeur { get; set; }
    }

    public class MesureOrchestrator
    {
        private readonly IIeeeDriver _driver;
        private readonly IExcelService _excel;

        public MesureOrchestrator(IIeeeDriver driver, IExcelService excel)
        {
            _driver = driver;
            _excel = excel;
        }

        /// <summary>Débloque la mesure GPIB en cours sur Arrêter. Cancel() ne suffit pas : le ReadString NI-VISA
        /// reste bloqué jusqu'à la fin de la gate (jusqu'à 1000 s), alors qu'un SDC le libère en quelques ms.</summary>
        public void AborterMesureEnCours() => _driver.AborterToutesSessions();

        public async Task<ResultatMesure> ExecuterAsync(
            Mesure mesure,
            Rubidium rubidium,
            double? fNominaleOuReference,
            IProgress<ProgressionMesure>? progress = null,
            CancellationToken ct = default)
        {
            var result = new ResultatMesure();

            // Profiling accumulé en mémoire (~10 ns par marqueur) : écrire chaque Perf() en SQL coûtait 5 à 10 ms,
            // soit jusqu'à 300 ms par gate sur les mesures rapides.
            var profiler = new ProfilerSession();
            void Perf(string label) => profiler.Mark(label);

            // Déclarés avant le try : le nettoyage d'annulation (catch/finally) en a besoin.
            string? derniereFeuille = null;
            // Vrai de la création de la feuille d'une gate à la fin de son acquisition : en cas d'arrêt,
            // indique si la feuille courante est incomplète (à supprimer) ou si la gate est terminée (à garder).
            bool feuilleCouranteIncomplete = false;
            // Chemin du .xlsx, pour supprimer la feuille sur disque après un stop (finally).
            string? cheminFichier = null;
            // Vrai si une valeur sort du domaine du module d'incertitude : la feuille courante est alors supprimée, comme pour un arrêt.
            bool horsModule = false;

            // Chronos réels pour EstimationDureeStab (apprend l'init de l'appareil et le surcoût de chaque temps de porte sur ce poste).
            var swInit = Stopwatch.StartNew();

            try
            {
                Perf("START");

                // 0. Remise à zéro des sessions GPIB en cache du driver VISA : évite les timeouts quand un appareil
                //    a été éteint puis rallumé entre deux mesures.
                _driver.ReinitialiserSessions();
                Perf("ReinitialiserSessions");

                // 1. Résolution de l'appareil dans le catalogue local.
                if (string.IsNullOrWhiteSpace(mesure.IdModeleCatalogue))
                {
                    result.Erreur = "Aucun appareil sélectionné pour la mesure. "
                        + "Enregistrez le fréquencemètre via Administration → Gérer les appareils.";
                    return result;
                }

                var (appareil, erreur) = ResolverDepuisCatalogue(mesure.IdModeleCatalogue, mesure);
                if (appareil == null)
                {
                    result.Erreur = erreur;
                    return result;
                }

                // 2. Initialisation et configuration de l'appareil (IFC, ConfEntree, SRQ on, puis rejeu SCPI).
                progress?.Report(new ProgressionMesure
                {
                    Message = $"Initialisation du {appareil.Nom} (GPIB {appareil.Adresse})...",
                });

                if (!mesure.InitManu)
                {
                    await appareil.InitialiserAsync(_driver, ct);
                    Perf("InitialiserAsync (*RST;*CLS)");
                }
                else
                {
                    // Init manuelle : pas de *RST/*CLS, qui écraseraient la config faite à la main par l'opérateur.
                    JournalLog.Info(CategorieLog.Mesure, "INIT_MANUELLE",
                        $"Init manuelle : *RST/*CLS et config non envoyés à {appareil.Nom} " +
                        "(configuration assurée à la main par l'opérateur).");
                }
                await appareil.ConfigurerAsync(_driver, mesure, mux: null, commandesMux: null, ct);
                Perf("ConfigurerAsync (IFC + ConfEntree + SRQ)");
                await RejouerScpiPostResetAsync(appareil, mesure, ct);
                Perf("RejouerScpiPostResetAsync");

                if (fNominaleOuReference.HasValue)
                {
                    mesure.FNominale = fNominaleOuReference.Value;
                    result.FNominale = fNominaleOuReference.Value;
                }

                // 3. Boucle de balayage : une itération en général, N pour une stabilité multi-gate.
                //    Chaque itération crée sa propre feuille Excel et sa ligne dans Récap.
                var gates = mesure.GateIndices.Count > 0
                    ? mesure.GateIndices.ToList()
                    : new List<int> { 6 };  // Liste vide : repli sur 1 s.

                int nbIterations = gates.Count;
                int totalEtapes = nbIterations * (mesure.NbMesures + 3);
                int etape = 0;

                bool bulkDejaEchoue = false;

                if (mesure.TypeMesure == TypeMesure.Stabilite)
                {
                    // Remise à zéro des sigmas accumulés (calibration de l'axe Y du graphe Stab à la dernière gate).
                    ExcelInteropHost.Instance.ReinitialiserSigmasStab();
                }

                for (int g = 0; g < nbIterations; g++)
                {
                    // Pas de gate suivante si l'utilisateur a stoppé entre deux gates (le ReadString de la gate
                    // courante peut avoir rendu la main avant le Device Clear).
                    ct.ThrowIfCancellationRequested();

                    bool isDernier = g == nbIterations - 1;
                    int gateIdx = gates[g];
                    Perf($"--- Gate {g + 1}/{nbIterations} ({EnTetesMesureHelper.LibelleGate(gateIdx)}) ---");
                    if (g == 0 && mesure.TypeMesure == TypeMesure.Stabilite)
                        EstimationDureeStab.EnregistrerInit(swInit.Elapsed);
                    var swGate = Stopwatch.StartNew();

                    progress?.Report(new ProgressionMesure
                    {
                        Message = nbIterations > 1
                            ? $"Gate {g + 1}/{nbIterations} ({EnTetesMesureHelper.LibelleGate(gateIdx)}) — préparation Excel..."
                            : $"Préparation du rapport Excel pour FI {mesure.NumFI}...",
                        EtapeActuelle = ++etape,
                        EtapesTotales = totalEtapes
                    });

                    // Le classeur est préparé avec ClosedXML (plus robuste que l'ajout de feuille en COM : pas
                    // d'erreur 0x800A03EC sur les formules métier), puis rouvert dans Excel pour l'utilisateur.
                    await ExcelInteropHost.Instance.FermerClasseurActifAsync();
                    Perf("Interop.FermerClasseurActif (avant init ClosedXML)");
                    // Excel COM relâche son handle de fichier de façon asynchrone : sans cette attente,
                    // FichierEstVerrouille (InitialiserRapportAsync) renvoie true à cause du cache MRU d'Excel.
                    await Task.Delay(500);

                    // 3.a Nouvelle feuille de mesure (1, 2, 3… en stab, selon le slot dispo). À la 1ère gate d'une
                    // session Stab, "nouvelle session" fait passer à un nouveau fichier (stab1, stab2…) si le précédent
                    // existe, pour que le graphe Stab ne reprenne pas les valeurs de la mesure précédente sur la même FI.
                    bool nouvelleSession = (g == 0 && mesure.TypeMesure == TypeMesure.Stabilite);
                    await _excel.InitialiserRapportAsync(mesure.NumFI, mesure, rubidium, gateIdx, nouvelleSession);
                    Perf("ClosedXML.InitialiserRapportAsync");
                    await _excel.PreparerLignesMesureAsync(mesure.NbMesures);
                    Perf("ClosedXML.PreparerLignesMesureAsync");

                    cheminFichier = await _excel.SauvegarderSurDisqueAsync();
                    Perf("ClosedXML.SauvegarderSurDisqueAsync (.xlsx)");
                    string nomFeuille = _excel.NomFeuilleMesure;
                    derniereFeuille = nomFeuille;

                    _excel.FermerExcel();
                    Perf("ClosedXML.FermerExcel");
                    await ExcelInteropHost.Instance.OuvrirEtAfficherAsync(cheminFichier, nomFeuille);
                    Perf("Interop.OuvrirEtAfficherAsync");

                    // Feuille créée, valeurs pas encore acquises : un stop maintenant la supprime (cf. catch
                    // OperationCanceledException). Repassé à false en fin de gate.
                    feuilleCouranteIncomplete = true;

                    // 3.b Programmation de la gate de cette itération.
                    progress?.Report(new ProgressionMesure
                    {
                        Message = nbIterations > 1
                            ? $"Gate {g + 1}/{nbIterations} ({EnTetesMesureHelper.LibelleGate(gateIdx)}) — programmation..."
                            : $"Programmation de la gate (index {gateIdx})...",
                        EtapeActuelle = ++etape,
                        EtapesTotales = totalEtapes
                    });

                    // Entre deux gates : :ABORT + *CLS pour repartir avec un buffer de sortie vide. Sans ça le 53131A
                    // (et la plupart des compteurs) garde l'état FETCH de la gate précédente et le :READ:FREQ? suivant
                    // part en timeout (constaté : la 1ère gate, juste après *RST, passe, les suivantes timeoutent).
                    if (g > 0)
                    {
                        try
                        {
                            await _driver.EcrireAsync(appareil.Adresse, ":ABORT", appareil.WriteTerm, ct);
                            await Task.Delay(20, ct);
                            await _driver.EcrireAsync(appareil.Adresse, "*CLS", appareil.WriteTerm, ct);
                            await Task.Delay(50, ct);
                            JournalLog.Info(CategorieLog.Mesure, "INTER_GATE_RESET",
                                $"Réinit entre gates : :ABORT + *CLS envoyés à GPIB0::{appareil.Adresse}");
                        }
                        catch (Exception ex)
                        {
                            JournalLog.Warn(CategorieLog.Mesure, "INTER_GATE_RESET_ERR",
                                $"Réinit entre gates échoué : {ex.Message} — la mesure continue.");
                        }
                        Perf("Inter-gate reset (:ABORT + *CLS)");
                    }

                    // Vérification d'arming sautée après la 1ère gate d'un balayage stab (déjà éprouvée, ~200 ms gagnées par gate).
                    bool verifierArming = !(mesure.TypeMesure == TypeMesure.Stabilite && g > 0);
                    await appareil.AppliquerGateAsync(_driver, gateIdx, mesure.TypeMesure, ct, verifierArming);
                    Perf($"AppliquerGateAsync (verifArming={verifierArming})");

                    // Intervalle de temps : le READ? bloque jusqu'à l'évènement et AppliquerGateAsync sort tôt (pas de gate),
                    // donc aucun timeout n'est posé et le défaut (~10 s) ferait échouer un intervalle long. L'intervalle
                    // pouvant aller jusqu'à 3600 s, on cale le timeout VISA dessus plus une marge. Les autres types gardent
                    // leur timeout basé sur la gate.
                    if (mesure.TypeMesure == TypeMesure.Interval)
                    {
                        const int timeoutIntervalleMs = (3600 + 300) * 1000;   // 3600 s max, plus 5 min de marge
                        _driver.DefinirTimeout(appareil.Adresse, timeoutIntervalleMs);
                        Perf($"Timeout intervalle = {timeoutIntervalleMs} ms");
                    }

                    // 3.c Boucle des N mesures.
                    //   - Stabilité, gate COURTE (< 500 ms) : accumulation en RAM, écriture en bloc en fin de gate
                    //     (chaque écriture COM live coûte ~300 ms de re-render : 22 s → 3-4 s typiquement).
                    //   - Stabilité, gate LONGUE (≥ 500 ms) : écriture en live, l'overhead est marginal.
                    //   - Autres types (Fréquence, Interval…) : toujours en live, à cadence humaine.
                    double gateSecondes = EnTetesMesureHelper.SecondesGate(gateIdx);
                    bool ecritureBatch = mesure.TypeMesure == TypeMesure.Stabilite
                                         && gateSecondes < 0.5;
                    var valeurs = new List<double>();
                    var bufferBatch = ecritureBatch ? new List<(DateTime, double)>(mesure.NbMesures) : null;
                    const int LIGNE_DEBUT_MESURES = 9;

                    // Trois stratégies d'acquisition, de la plus rapide à la plus standard :
                    //   1. Bulk    : CommandeMesureMultiple (avec {N}) : l'instrument fait les N mesures et les renvoie
                    //                en CSV (~0,5 s pour 30 mesures à 10 ms, contre ~6 s en :FETCh?).
                    //   2. Rapide  : :INIT:CONT ON puis boucle :FETCh?, sans ré-arming du :READ? (~180 ms/mesure contre 670 ms).
                    //   3. Classic : un :READ? par mesure (ré-arme à chaque appel).
                    // Le choix vient du catalogue (CommandeMesureMultiple) : aucune logique propre à un modèle d'appareil ici.
                    string? cmdFetch = AppareilIeeeOperations.DeriverCommandeFetch(appareil.ExeMesure);

                    // Commande « fetch fresh » renseignée (bloque jusqu'à une nouvelle mesure) : prioritaire, sans
                    // Task.Delay, c'est l'instrument qui donne la cadence.
                    bool fetchBloquant = !string.IsNullOrWhiteSpace(appareil.CommandeFetchFresh);
                    string? cmdFetchUsed = fetchBloquant ? appareil.CommandeFetchFresh : cmdFetch;

                    // NbMesures == 1 : ni streaming ni mode rapide (l'init bulk coûte plus cher que la mesure), :READ? classique.
                    // Les garde-fous ci-dessous (CommandeBulkInit/FetchFresh du catalogue, gate ≥ 1 s, bulkDejaEchoue)
                    // écartent de toute façon les instruments qui ne supportent pas le streaming.
                    bool typeFaitSeries = mesure.NbMesures > 1;

                    // Mode streaming : acquisition gap-free lue une valeur à la fois via CommandeFetchFresh
                    // (typiquement ":DATA:REM? 1,WAIT" sur 53230A), compatible UI live et batch Stab. Réservé aux gates ≥ 1 s :
                    // en dessous, chaque DATA:REM coûte 30 à 50 ms de roundtrip GPIB qui polluerait la cadence (bulk classique).
                    bool modeStreamingPossible = !string.IsNullOrWhiteSpace(appareil.CommandeBulkInit)
                        && !string.IsNullOrWhiteSpace(appareil.CommandeFetchFresh)
                        && gateSecondes >= 1.0
                        && typeFaitSeries
                        && !bulkDejaEchoue;

                    // Mode bulk : timestamps simulés (l'instrument mesure en interne), incompatible avec l'écriture live
                    // Excel cellule par cellule ; réservé à la Stabilité, qui bufferise de toute façon. Exclusif du streaming.
                    bool modeBulk = !modeStreamingPossible
                        && !string.IsNullOrWhiteSpace(appareil.CommandeMesureMultiple)
                        && mesure.TypeMesure == TypeMesure.Stabilite
                        && !bulkDejaEchoue;
                    // Gates ≥ 1 s : un bulk en un coup ne rendrait les N valeurs qu'à la fin (50 min à 100 s × 30). On garde la
                    // MÊME commande avec {N}=1, une transaction par mesure, pour voir chaque valeur dans Excel à son arrivée
                    // (choix utilisateur, 1 s incluse). Surtout pas de bascule sur :INIT:CONT ON + :FETCh? : -113 Undefined
                    // header sur les 53230A et consorts.
                    bool bulkUnitaire = modeBulk && gateSecondes >= 1.0;
                    // Mode rapide :FETCh? : tout type avec NbMesures > 1 si le catalogue l'autorise (ModeRapideActif),
                    // compatible écriture live Excel, ~3× plus rapide que :READ? sur le 53131A.
                    bool modeRapide = !modeBulk
                        && !modeStreamingPossible
                        && !string.IsNullOrEmpty(cmdFetchUsed)
                        && typeFaitSeries
                        && appareil.ModeRapideActif;   // bloqué pour le 53230A et consorts (cf. catalogue)

                    // Délai entre deux fetch : 0 si la commande est bloquante (l'instrument attend), sinon gate - 100 ms
                    // (plancher 5 ms). Le write+read GPIB prend ~150 ms, donc Delay + 150 ≥ gate garantit une nouvelle
                    // mesure disponible au :FETCh?.
                    int delayFetchMs;
                    if (!modeRapide) delayFetchMs = 0;
                    else if (fetchBloquant) delayFetchMs = 0;
                    else delayFetchMs = Math.Max(5, (int)(gateSecondes * 1000) - 100);

                    if (modeRapide)
                    {
                        try
                        {
                            await _driver.EcrireAsync(appareil.Adresse, ":INIT:CONT ON", appareil.WriteTerm, ct);
                            // Amorçage adaptatif : délai = gate + 200 ms (mini 150 ms) pour que la 1ère mesure interne soit
                            // complète. Si le fetch renvoie 0 (erreur GPIB 230 "Data stale"), retry avec délai ×2, ×3, ×4.
                            int amorceInitial = Math.Max(150, (int)(gateSecondes * 1000) + 200);
                            await Task.Delay(amorceInitial, ct);
                            for (int essai = 0; essai < 3; essai++)
                            {
                                var testVal = await appareil.FetcherAsync(_driver, cmdFetchUsed!, ct);
                                if (testVal != 0.0) break;
                                await Task.Delay(amorceInitial * (essai + 2), ct);
                            }
                            Perf($"INIT:CONT ON (mode rapide, fetch={cmdFetchUsed}, amorce={amorceInitial}ms, delay-loop={delayFetchMs}ms)");
                        }
                        catch (Exception ex)
                        {
                            JournalLog.Warn(CategorieLog.Mesure, "INIT_CONT_ON_ERR",
                                $"Activation mode rapide échouée : {ex.Message} — fallback :READ?.");
                            modeRapide = false;
                        }
                    }

                    var swBoucle = Stopwatch.StartNew();

                    // Écritures Excel Interop COM en fire-and-forget : chaque EcrireValeurLiveAsync prend ~50-100 ms (lock COM),
                    // on les recouvre avec la mesure GPIB suivante pour retrouver la cadence physique de l'instrument.
                    // Le lock(_sync) d'ExcelInteropHost sérialise les écritures en FIFO (pas de race à ≤ 10 Hz).
                    // Vide en mode batch (Stab) : pas d'Interop pendant les gates.
                    var pendingWrites = new List<Task>(mesure.NbMesures);

                    // ===== MODE STREAMING : prioritaire sur les autres modes si applicable =====
                    if (modeStreamingPossible)
                    {
                        // Timeout : gate × N + 3 s, comme le bulk classique.
                        int timeoutStream = Math.Max(5000,
                            (int)(gateSecondes * 1000 * mesure.NbMesures) + 3000 + 100 * mesure.NbMesures);
                        _driver.DefinirTimeout(appareil.Adresse, timeoutStream);

                        var swStream = Stopwatch.StartNew();
                        try
                        {
                            // 1. Lance l'acquisition gap-free en arrière-plan.
                            string cmdInit = appareil.CommandeBulkInit.Replace("{N}",
                                mesure.NbMesures.ToString(CultureInfo.InvariantCulture));
                            await _driver.EcrireAsync(appareil.Adresse, cmdInit, appareil.WriteTerm, ct);

                            // 2. Lecture de chaque valeur dès qu'elle est dispo (WAIT bloque jusqu'à sa présence en mémoire du compteur).
                            for (int i = 0; i < mesure.NbMesures; i++)
                            {
                                ct.ThrowIfCancellationRequested();
                                double val = await appareil.FetcherAsync(
                                    _driver, appareil.CommandeFetchFresh, ct);
                                var ts = DateTime.Now;
                                valeurs.Add(val);

                                if (ecritureBatch)
                                    bufferBatch!.Add((ts, val));
                                else
                                    pendingWrites.Add(ExcelInteropHost.Instance.EcrireValeurLiveAsync(i, val, ts));

                                // Throttle UI : une mesure sur 5, plus la dernière.
                                if ((i % 5 == 0) || (i == mesure.NbMesures - 1))
                                {
                                    progress?.Report(new ProgressionMesure
                                    {
                                        Message = nbIterations > 1
                                            ? $"Gate {g + 1}/{nbIterations} — mesure {i + 1}/{mesure.NbMesures}"
                                            : $"Mesure {i + 1}/{mesure.NbMesures}",
                                        EtapeActuelle = etape + i + 1,
                                        EtapesTotales = totalEtapes,
                                        DerniereValeur = val
                                    });
                                }
                            }
                            swStream.Stop();
                            JournalLog.Info(CategorieLog.Mesure, "BULK_STREAM",
                                $"Streaming gap-free : {mesure.NbMesures} valeurs lues en "
                              + $"{swStream.ElapsedMilliseconds} ms via DATA:REM (live UI).");
                            etape += mesure.NbMesures;
                            goto FinBoucleMesures;
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception ex)
                        {
                            JournalLog.Warn(CategorieLog.Mesure, "BULK_STREAM_ERR",
                                $"Mode streaming échoué : {ex.Message} — fallback sur les autres modes.");
                            try { await _driver.EcrireAsync(appareil.Adresse, ":ABORT", appareil.WriteTerm, ct); }
                            catch { /* best-effort */ }
                            valeurs.Clear();
                            pendingWrites.Clear();
                            bufferBatch?.Clear();
                        }
                    }

                    if (modeBulk)
                    {
                        // ----- Mode BULK : 1 seul aller-retour GPIB pour N mesures -----
                        // Timeout : gate × N + 3 s de marge (init mode CONT, transmission CSV, parsing)
                        // + 100 ms par mesure de temps mort entre portes (une marge fixe devenait juste à 100 mesures).
                        int timeoutBulk = Math.Max(5000,
                            (int)(gateSecondes * 1000 * mesure.NbMesures) + 3000 + 100 * mesure.NbMesures);
                        _driver.DefinirTimeout(appareil.Adresse, timeoutBulk);

                        // En bulk, les N mesures sont faites en interne et ne remontent pas une par une à l'utilisateur.
                        double dureeAttendueSec = gateSecondes * mesure.NbMesures;
                        if (!bulkUnitaire)   // en unitaire, chaque mesure fait son propre report
                        progress?.Report(new ProgressionMesure
                        {
                            Message = nbIterations > 1
                                ? $"Gate {g + 1}/{nbIterations} ({EnTetesMesureHelper.LibelleGate(gateIdx)}) — "
                                  + $"acquisition gap-free de {mesure.NbMesures} mesures en cours (~{dureeAttendueSec:F1} s)..."
                                : $"Acquisition gap-free de {mesure.NbMesures} mesures en cours (~{dureeAttendueSec:F1} s)...",
                            EtapeActuelle = etape,
                            EtapesTotales = totalEtapes
                        });

                        bool bulkOk = false;
                        var swBulkProgress = Stopwatch.StartNew();

                        if (bulkUnitaire)
                        {
                            // Une mesure par transaction : timeout calé sur UNE gate (+5 s de marge).
                            _driver.DefinirTimeout(appareil.Adresse,
                                Math.Max(5000, (int)(gateSecondes * 1000) + 5000));
                            try
                            {
                                for (int i = 0; i < mesure.NbMesures; i++)
                                {
                                    ct.ThrowIfCancellationRequested();
                                    var lot1 = await appareil.MesurerEnLotAsync(
                                        _driver, appareil.CommandeMesureMultiple, 1, ct);
                                    if (lot1.Count < 1)
                                        throw new InvalidOperationException(
                                            $"réponse vide à la mesure {i + 1}/{mesure.NbMesures}");

                                    double val = lot1[0];
                                    var ts = DateTime.Now;
                                    valeurs.Add(val);
                                    if (ecritureBatch)
                                        bufferBatch!.Add((ts, val));
                                    else
                                        pendingWrites.Add(ExcelInteropHost.Instance.EcrireValeurLiveAsync(i, val, ts));

                                    // Gate ≥ 1 s : un report par mesure ne sature pas l'UI.
                                    progress?.Report(new ProgressionMesure
                                    {
                                        Message = nbIterations > 1
                                            ? $"Gate {g + 1}/{nbIterations} — mesure {i + 1}/{mesure.NbMesures}"
                                            : $"Mesure {i + 1}/{mesure.NbMesures}",
                                        EtapeActuelle = etape + i + 1,
                                        EtapesTotales = totalEtapes,
                                        DerniereValeur = val
                                    });
                                }
                                bulkOk = true;
                                JournalLog.Info(CategorieLog.Mesure, "BULK_UNITAIRE",
                                    $"{mesure.NbMesures} mesures lues une par une ({{N}}=1), gate {gateSecondes} s.");
                            }
                            catch (OperationCanceledException) { throw; }
                            catch (Exception ex)
                            {
                                JournalLog.Warn(CategorieLog.Mesure, "BULK_UNITAIRE_ERR",
                                    $"Mesure une par une échouée : {ex.Message} — fallback.");
                                valeurs.Clear();
                                bufferBatch?.Clear();
                            }
                        }
                        else
                        try
                        {
                            var lot = await appareil.MesurerEnLotAsync(
                                _driver, appareil.CommandeMesureMultiple, mesure.NbMesures, ct);
                            int nbRecu = lot.Count;
                            swBulkProgress.Stop();

                            // Bulk OK si au moins 80 % des mesures reçues ; sinon (commande refusée, retour vide,
                            // erreur SCPI) bascule sur le mode :FETCh? rapide.
                            if (nbRecu >= mesure.NbMesures * 0.8)
                            {
                                JournalLog.Info(CategorieLog.Mesure, "BULK_FETCH",
                                    $"Lot reçu : {nbRecu}/{mesure.NbMesures} valeurs en 1 transaction GPIB.");

                                progress?.Report(new ProgressionMesure
                                {
                                    Message = nbIterations > 1
                                        ? $"Gate {g + 1}/{nbIterations} — {nbRecu} mesures reçues en {swBulkProgress.ElapsedMilliseconds / 1000.0:F1} s"
                                        : $"{nbRecu} mesures reçues en {swBulkProgress.ElapsedMilliseconds / 1000.0:F1} s",
                                    EtapeActuelle = etape,
                                    EtapesTotales = totalEtapes
                                });

                                var tsBase = DateTime.Now.AddMilliseconds(-(gateSecondes * 1000) * mesure.NbMesures);
                                for (int i = 0; i < mesure.NbMesures; i++)
                                {
                                    double val = i < nbRecu ? lot[i] : 0.0;
                                    var ts = tsBase.AddMilliseconds(gateSecondes * 1000 * i);
                                    valeurs.Add(val);
                                    if (ecritureBatch)
                                        bufferBatch!.Add((ts, val));
                                    else
                                        pendingWrites.Add(ExcelInteropHost.Instance.EcrireValeurLiveAsync(i, val, ts));
                                }
                                bulkOk = true;
                            }
                            else
                            {
                                JournalLog.Warn(CategorieLog.Mesure, "BULK_FETCH_INSUFFISANT",
                                    $"Bulk a renvoyé seulement {nbRecu}/{mesure.NbMesures} valeurs — la commande "
                                    + "n'est probablement pas adaptée. Fallback sur :FETCh? boucle.");
                            }
                        }
                        catch (Exception ex)
                        {
                            JournalLog.Warn(CategorieLog.Mesure, "BULK_FETCH_ERR",
                                $"Mode bulk échoué : {ex.Message} — fallback sur :FETCh? boucle.");
                        }

                        if (bulkOk) goto FinBoucleMesures;

                        // Plus de bulk pour les gates suivantes.
                        bulkDejaEchoue = true;
                        JournalLog.Info(CategorieLog.Mesure, "BULK_DESACTIVE_POUR_BALAYAGE",
                            "Commande bulk désactivée pour le reste du balayage — la 1ère tentative a échoué.");

                        // Fallback léger : :ABORT + *CLS purgent une éventuelle query interrompue. Pas de *RST :
                        // l'instrument n'est pas corrompu, les réglages SCPI utilisateur restent en place.
                        try
                        {
                            await _driver.EcrireAsync(appareil.Adresse, ":ABORT", appareil.WriteTerm, ct);
                            await Task.Delay(20, ct);
                            await _driver.EcrireAsync(appareil.Adresse, "*CLS", appareil.WriteTerm, ct);
                            await Task.Delay(50, ct);
                        }
                        catch { /* best effort */ }

                        // Timeout VISA normal rétabli (basé sur la gate, comme dans AppliquerGateAsync).
                        int timeoutNormal = Math.Max(5000, (int)(gateSecondes * 1000) + 2000);
                        _driver.DefinirTimeout(appareil.Adresse, timeoutNormal);

                        modeBulk = false;
                        // ModeRapideActif = false sur les compteurs modernes (:INIT:CONT ON + :FETCh? donne -113) : :READ? classique.
                        modeRapide = !string.IsNullOrEmpty(cmdFetch) && appareil.ModeRapideActif;
                        if (modeRapide)
                        {
                            // delayFetchMs était à 0 (modeRapide faux avant le fallback).
                            delayFetchMs = Math.Max(15, (int)(gateSecondes * 1000) + 5);
                            try
                            {
                                await _driver.EcrireAsync(appareil.Adresse, ":INIT:CONT ON", appareil.WriteTerm, ct);
                                await Task.Delay(delayFetchMs + 50, ct);
                            }
                            catch { modeRapide = false; }
                        }
                    }

                    // Timeouts FetchFresh : si la commande « fresh » renvoie 0 (timeout VISA + Device Clear) 3 fois
                    // d'affilée, l'instrument ne la reconnaît pas : bascule en :FETCh? classique avec Task.Delay.
                    int freshTimeoutsConsec = 0;

                    // Legacy à SRQ (EIP 545...) : la config (SR01) et le réglage de gate peuvent laisser un SRQ « périmé »
                    // qui ferait renvoyer une valeur parasite à la 1ère lecture (« 1ère valeur qui saute »). Un serial poll
                    // de purge vide ce statut (no-op sans SRQ ou en mode rapide).
                    if (appareil.GereSRQ && !modeRapide)
                    {
                        try { await _driver.LireStatusByteAsync(appareil.Adresse, ct); }
                        catch (OperationCanceledException) { throw; }
                        catch { /* best-effort : la purge n'est pas critique */ }
                    }

                    for (int i = 0; i < mesure.NbMesures; i++)
                    {
                        ct.ThrowIfCancellationRequested();

                        double val;
                        if (modeRapide)
                        {
                            if (i > 0 && delayFetchMs > 0) await Task.Delay(delayFetchMs, ct);
                            val = await appareil.FetcherAsync(_driver, cmdFetchUsed!, ct);

                            if (fetchBloquant && val == 0.0)
                            {
                                freshTimeoutsConsec++;
                                if (freshTimeoutsConsec >= 3)
                                {
                                    JournalLog.Warn(CategorieLog.Mesure, "FETCH_FRESH_DESACTIVE",
                                        $"FetchFresh « {appareil.CommandeFetchFresh} » timeout 3× — "
                                        + "fallback sur :FETCh? + Task.Delay pour le reste de la gate.");
                                    fetchBloquant = false;
                                    cmdFetchUsed = cmdFetch;
                                    delayFetchMs = Math.Max(5, (int)(gateSecondes * 1000 * 0.5));
                                    // Ré-arme le mode continu (FetchFresh a pu le perturber).
                                    try
                                    {
                                        await _driver.EcrireAsync(appareil.Adresse, ":ABORT", appareil.WriteTerm, ct);
                                        await Task.Delay(20, ct);
                                        await _driver.EcrireAsync(appareil.Adresse, "*CLS", appareil.WriteTerm, ct);
                                        await Task.Delay(50, ct);
                                        await _driver.EcrireAsync(appareil.Adresse, ":INIT:CONT ON", appareil.WriteTerm, ct);
                                        await Task.Delay(delayFetchMs + 50, ct);
                                    }
                                    catch { /* best-effort */ }
                                    i--;  // retry cette mesure en mode classique
                                    continue;
                                }
                            }
                            else
                            {
                                freshTimeoutsConsec = 0;
                            }

                            // Pas de détection de doublons : deux valeurs identiques sur une source stable sont valides, et
                            // delayFetchMs garantit déjà la lecture d'une gate complète.
                        }
                        else
                        {
                            // Durée de gate (ms) : MesurerAsync attend le MAV si l'appareil le gère ; sinon (legacy type EIP
                            // sans MAV via VISA) il cale un délai sur la gate, faute de quoi la lecture part trop tôt (0 en rafale).
                            int gateMs = (int)(gateSecondes * 1000);
                            val = await appareil.MesurerAsync(_driver, ct, gateMs);
                        }
                        var ts = DateTime.Now;
                        valeurs.Add(val);

                        if (ecritureBatch)
                        {
                            bufferBatch!.Add((ts, val));
                        }
                        else
                        {
                            // Fire-and-forget, en parallèle de la prochaine mesure GPIB ; flushé en fin de boucle (Task.WhenAll).
                            pendingWrites.Add(ExcelInteropHost.Instance.EcrireValeurLiveAsync(i, val, ts));
                        }

                        // Throttle UI (Stab + Freq) : une mesure sur 5, plus la dernière. Chaque Report() re-rend le TextBox
                        // sur le thread dispatcher (200-500 ms avec InformationsGenerales de plusieurs Ko) : sans throttle,
                        // les 30 reports/gate saturaient l'UI et le bouton « Arrêter » devenait incliquable.
                        bool reporter = (i % 5 == 0) || (i == mesure.NbMesures - 1);
                        if (reporter)
                        {
                            progress?.Report(new ProgressionMesure
                            {
                                Message = nbIterations > 1
                                    ? $"Gate {g + 1}/{nbIterations} — mesure {i + 1}/{mesure.NbMesures}"
                                    : $"Mesure {i + 1}/{mesure.NbMesures}",
                                EtapeActuelle = etape + i + 1,
                                EtapesTotales = totalEtapes,
                                DerniereValeur = val
                            });
                        }
                    }
                FinBoucleMesures:
                    etape += mesure.NbMesures;

                    // Flush des écritures Excel en vol : la boucle GPIB peut finir avant les dernières cellules COM,
                    // et la finalisation partirait sur des cellules encore vides.
                    if (pendingWrites.Count > 0)
                    {
                        try { await Task.WhenAll(pendingWrites); }
                        catch (Exception ex)
                        {
                            JournalLog.Warn(CategorieLog.Excel, "EXCEL_FLUSH_ERREUR",
                                $"Erreur durant flush des écritures live ({pendingWrites.Count} en attente) : {ex.Message}");
                        }
                        Perf($"Flush écritures Excel ({pendingWrites.Count} cellules)");
                    }

                    swBoucle.Stop();
                    string modeTag = modeBulk ? "BULK"
                        : modeRapide ? (fetchBloquant ? "FETCH-FRESH" : "FETCH")
                        : "READ";
                    Perf($"Boucle GPIB {mesure.NbMesures} mesures ({swBoucle.ElapsedMilliseconds} ms total, ~{swBoucle.ElapsedMilliseconds / Math.Max(1, mesure.NbMesures)} ms/mesure — {modeTag})");

                    // Mode continu coupé pour que l'instrument ne trigge pas en arrière-plan entre les gates ou après la mesure.
                    if (modeRapide)
                    {
                        try
                        {
                            await _driver.EcrireAsync(appareil.Adresse, ":INIT:CONT OFF", appareil.WriteTerm, ct);
                            Perf("INIT:CONT OFF (fin mode rapide)");
                        }
                        catch (Exception ex)
                        {
                            JournalLog.Warn(CategorieLog.Mesure, "INIT_CONT_OFF_ERR",
                                $"Désactivation mode rapide échouée : {ex.Message}");
                        }
                    }

                    // Flush batch des N valeurs accumulées via Interop (Range.Value2 = matrice).
                    if (ecritureBatch && bufferBatch!.Count > 0)
                    {
                        await ExcelInteropHost.Instance
                            .EcrireValeursEnBlocAsync(LIGNE_DEBUT_MESURES, bufferBatch);
                        Perf("Interop.EcrireValeursEnBlocAsync (bulk write)");
                    }

                    // 3.d SRQ off uniquement à la dernière itération, pour ne pas le désactiver/réactiver entre
                    //     deux gates (cf. F_Main:1262, correctif Racal historique).
                    if (isDernier)
                    {
                        await appareil.DesactiverSrqAsync(_driver, ct);
                        Perf("DesactiverSrqAsync");
                    }

                    // 3.e Stats + ligne Récap pour cette gate, tout via COM dans le Workbook ouvert (pas de cycle
                    //     ferme/rouvre, donc pas de fenêtre Frame grise en fin de mesure).
                    progress?.Report(new ProgressionMesure
                    {
                        Message = nbIterations > 1
                            ? $"Gate {g + 1}/{nbIterations} — finalisation..."
                            : "Finalisation du rapport Excel...",
                        EtapeActuelle = etape,
                        EtapesTotales = totalEtapes
                    });

                    // Garde-fou « dépassement de module » : moyenne hors du domaine du module d'incertitude sélectionné
                    // → exception typée ; le finally supprime la feuille en cours (comme un arrêt), le ViewModel affiche la popup.
                    // Module introuvable ≠ hors plage : le fallback historique est conservé.
                    var couvertureGate = IncertitudeCouverture.Verifier(mesure, valeurs, gateSecondes);
                    if (couvertureGate.EstHorsPlage)
                    {
                        string msgUtilisateur =
                            $"La valeur mesurée ({couvertureGate.ValeurLookup:G6} {couvertureGate.Unite}) "
                          + $"dépasse le domaine couvert par le module d'incertitude « {mesure.NumModuleIncertitude} ».\n\n"
                          + "La feuille de cette mesure a été supprimée.\n\n"
                          + "Vérifiez le module sélectionné, la fréquence nominale ou le temps de porte.";
                        string msgLog =
                            $"Valeur {couvertureGate.ValeurLookup:G6} {couvertureGate.Unite} hors plage du module "
                          + $"{mesure.NumModuleIncertitude} (gate={gateSecondes}s) — feuille « {nomFeuille} » supprimée.";
                        throw new MesureHorsModuleException(msgUtilisateur, msgLog);
                    }

                    // Stab : à la dernière gate le graphe est aussi patché via COM (plages des séries + bornes
                    // de l'axe Y log). Le dossier FI est transféré en bloc en fin de mesure
                    // (TransfertReseauService.TransfererDossierFIAsync).
                    int nbGatesStab = mesure.TypeMesure == TypeMesure.Stabilite ? (g + 1) : 0;
                    await ExcelInteropHost.Instance.FinaliserMesureViaComAsync(
                        mesure, valeurs, nomFeuille, gateSecondes, isDernier, nbGatesStab);
                    Perf("Interop.FinaliserMesureViaComAsync");

                    // Le graphe de stabilité (Écart type / Maxi / Mini en fonction du temps de porte) est embarqué dans le
                    // template METROLOGO_Stab.xlsx, extrait de Stab1.xls avec ses propriétés exactes (HiLoLines, Y log,
                    // markers, légende) : il est cloné avec chaque nouveau classeur, jamais recréé par code.

                    result.Valeurs.AddRange(valeurs);

                    JournalLog.Info(CategorieLog.Mesure, "Iteration",
                        $"Gate {g + 1}/{nbIterations} terminée : {valeurs.Count} valeurs sur {appareil.Nom}.",
                        new
                        {
                            appareil.Nom,
                            appareil.Adresse,
                            GateIndex = gateIdx,
                            GateLibelle = EnTetesMesureHelper.LibelleGate(gateIdx),
                            NbMesures = valeurs.Count
                        });

                    // Gate terminée : un arrêt pendant la suivante ne doit pas supprimer cette feuille.
                    feuilleCouranteIncomplete = false;

                    if (mesure.TypeMesure == TypeMesure.Stabilite)
                        EstimationDureeStab.EnregistrerGate(gateIdx, mesure.NbMesures, swGate.Elapsed);
                }

                if (cheminFichier != null) result.CheminExcel = cheminFichier;

                JournalLog.Info(CategorieLog.Mesure, "Execute",
                    $"Mesure terminée : {nbIterations} gate(s), {result.Valeurs.Count} valeur(s) au total.",
                    new
                    {
                        appareil.Nom,
                        NbGates = nbIterations,
                        NbValeursTotal = result.Valeurs.Count,
                        Gates = gates
                    });

                // Conversion .xls (réglage Admin) AVANT le transfert, pour que le réseau reçoive le format final.
                await ConvertirRapportSiXlsAsync(mesure, cheminFichier, result);
                Perf("ConvertirRapportSiXlsAsync");

                // Transfert du dossier FI complet vers le partage réseau M:\ (snapshot en bloc). Si KO (M:\ down,
                // latence…), la FI est inscrite en liste de reprise pour le prochain démarrage et le ViewModel avertit l'utilisateur.
                if (!string.IsNullOrWhiteSpace(mesure.NumFI))
                {
                    result.TransfertReseauOk = await TransfertReseauService
                        .TransfererDossierFIAsync(mesure.NumFI);
                    Perf("TransfertReseauService.TransfererDossierFIAsync");
                }

                result.Succes = true;
            }
            catch (OperationCanceledException)
            {
                result.Erreur = "Mesure annulée par l'utilisateur.";
                JournalLog.Warn(CategorieLog.Mesure, "Execute", result.Erreur);
            }
            catch (MesureHorsModuleException ex)
            {
                // Hors domaine du module : feuille supprimée (cf. finally), popup côté ViewModel via result.Erreur.
                horsModule = true;
                result.Erreur = ex.MessageUtilisateur;
                JournalLog.Warn(CategorieLog.Mesure, "MESURE_HORS_MODULE", ex.Message);
            }
            catch (Exception ex)
            {
                result.Erreur = ex.Message;
                JournalLog.Erreur(CategorieLog.Mesure, "Execute",
                    $"Échec de la mesure : {ex.Message}", new { ex.GetType().Name, ex.StackTrace });
            }
            finally
            {
                // Mesure stoppée : la feuille en cours est incomplète et doit être supprimée dès que l'annulation est
                // demandée, quelle que soit l'exception remontée (un Device Clear interrompt un :READ? bloquant par une
                // exception VISA, pas par OperationCanceledException). Uniquement si la feuille courante est incomplète,
                // sinon on effacerait une gate précédente déjà terminée.
                if ((ct.IsCancellationRequested || horsModule)
                    && feuilleCouranteIncomplete
                    && !string.IsNullOrEmpty(derniereFeuille))
                {
                    // 1. Tentative COM (utile si Excel est encore vivant).
                    try { await ExcelInteropHost.Instance.SupprimerFeuilleMesureAsync(derniereFeuille); }
                    catch (Exception exSuppr)
                    {
                        JournalLog.Warn(CategorieLog.Mesure, "Execute_SupprFeuilleStop",
                            $"Suppression COM de la feuille « {derniereFeuille} » après arrêt échouée : {exSuppr.Message}");
                    }

                    // 2. Suppression sur disque via ClosedXML : STOP tue le process Excel (TuerProcessExcelAsync), la
                    //    suppression COM est alors un no-op ; le .xlsx contient déjà la feuille sauvegardée. Chemin
                    //    réellement efficace après un STOP.
                    if (!string.IsNullOrEmpty(cheminFichier))
                    {
                        try { _excel.SupprimerFeuilleSurDisque(cheminFichier!, derniereFeuille!); }
                        catch (Exception exDisque)
                        {
                            JournalLog.Warn(CategorieLog.Mesure, "Execute_SupprFeuilleStopDisque",
                                $"Suppression sur disque de « {derniereFeuille} » après arrêt échouée : {exDisque.Message}");
                        }
                    }
                }

                Perf("Fin de mesure");
                _excel.FermerExcel();

                // Mesure arrêtée ou en échec : conversion .xls ici, la fin normale n'ayant pas été atteinte.
                if (!result.Succes && result.ConversionXlsOk == null)
                {
                    try { await ConvertirRapportSiXlsAsync(mesure, cheminFichier, result); }
                    catch (Exception exConv)
                    {
                        JournalLog.Warn(CategorieLog.Excel, "RAPPORT_CONVERSION_KO",
                            $"Conversion .xls après arrêt/échec impossible : {exConv.Message}");
                    }
                }

                // Le fichier Profiling_<FI>_<date>.txt n'est plus écrit (il polluait les dossiers de mesure) ;
                // le profiler reste alimenté en mémoire par Perf(...).
            }

            return result;
        }

        /// <summary>Réglage Admin « rapports en .xls » : convertit le rapport (hors stabilité, encore manuelle)
        /// en .xls et supprime le .xlsx. Sans effet en mode .xlsx.</summary>
        private static async Task ConvertirRapportSiXlsAsync(Mesure mesure, string? cheminFichier, ResultatMesure result)
        {
            if (mesure.TypeMesure == TypeMesure.Stabilite) return;
            if (Preferences.FormatRapportsActif != FormatRapport.Xls) return;
            if (string.IsNullOrEmpty(cheminFichier)
                || !cheminFichier.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                || !File.Exists(cheminFichier))
                return;

            result.ConversionXlsOk = await ExcelInteropHost.Instance.ConvertirRapportEnXlsAsync(cheminFichier);
            if (result.ConversionXlsOk == true)
                result.CheminExcel = Path.ChangeExtension(cheminFichier, ".xls");
        }

        /// <summary>Rejoue les commandes SCPI des réglages choisis dans Configuration (le <c>*RST</c> de la ChaineInit
        /// les efface). Un timeout est journalisé sans interrompre les suivantes ; 50 ms entre commandes pour les
        /// appareils lents (53131A).</summary>
        private async Task RejouerScpiPostResetAsync(AppareilIEEE appareil, Mesure mesure, CancellationToken ct)
        {
            if (mesure.CommandesScpiReglages == null || mesure.CommandesScpiReglages.Count == 0)
                return;

            foreach (var cmd in mesure.CommandesScpiReglages)
            {
                if (string.IsNullOrWhiteSpace(cmd)) continue;

                try
                {
                    await _driver.EcrireAsync(appareil.Adresse, cmd, appareil.WriteTerm, ct);
                    JournalLog.Info(CategorieLog.Mesure, "SCPI_REJEU",
                        $"GPIB0::{appareil.Adresse} ← {cmd} (réapplication post-RST)");
                }
                // VISA lève Ivi.Visa.IOTimeoutException, NI-488.2 une IOException contenant « timeout » : même traitement.
                catch (Exception ex) when (
                    ex is Ivi.Visa.IOTimeoutException ||
                    (ex is IOException && ex.Message.Contains("timeout", StringComparison.OrdinalIgnoreCase)))
                {
                    _ = ex; // utilisé dans le filtre
                    JournalLog.Warn(CategorieLog.Mesure, "SCPI_REJEU_TIMEOUT",
                        $"Timeout sur l'envoi de « {cmd} » à GPIB0::{appareil.Adresse} — "
                        + "commande ignorée, la mesure continue.");
                }

                await Task.Delay(50, ct);
            }
        }

        /// <summary>Retrouve le modèle catalogue et l'adresse GPIB détectée, et construit un <see cref="AppareilIEEE"/>
        /// via <see cref="CatalogueAdapter"/>.</summary>
        private static (AppareilIEEE?, string?) ResolverDepuisCatalogue(string idModele, Mesure mesure)
        {
            var modele = CatalogueAppareilsService.Instance.Modeles
                .FirstOrDefault(m => m.Id == idModele);
            if (modele == null)
            {
                return (null, $"Le modèle catalogue « {idModele} » est introuvable. "
                    + "Le catalogue a peut-être été modifié — rouvrez la Configuration.");
            }

            // Adresse fixe : forcée sur la mesure, ou appareil legacy (sans *IDN?) avec son adresse par défaut.
            // La détection sur le bus est court-circuitée (ces appareils n'apparaissent jamais dans AppareilsDetectes).
            int adresseFixe = mesure.AdresseFixeForcee >= 0
                ? mesure.AdresseFixeForcee
                : (modele.Parametres.Legacy ? modele.Parametres.AdresseFixeParDefaut : -1);

            if (adresseFixe >= 0)
                return (CatalogueAdapter.VersAppareilIEEE(modele, adresseFixe), null);

            // Appareil moderne détecté par *IDN? sur le bus.
            var detecte = EtatApplication.AppareilsDetectes
                .FirstOrDefault(d => modele.Correspond(d.Fabricant, d.Modele));
            if (detecte == null)
            {
                return (null, $"Le modèle « {modele.Nom} » n'est pas détecté sur le bus GPIB. "
                    + "Lancez un scan depuis Diagnostic GPIB pour le retrouver.");
            }

            return (CatalogueAdapter.VersAppareilIEEE(modele, detecte.Adresse), null);
        }
    }
}
