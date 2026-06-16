using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
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

        /// <summary>
        /// Transfert du dossier FI vers M:\ en fin de mesure : null = non tenté, true = ok,
        /// false = échec (FI mise en liste de reprise pour le prochain démarrage).
        /// </summary>
        public bool? TransfertReseauOk { get; set; }

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

        /// <summary>
        /// Débloque la mesure GPIB en cours (bouton Arrêter) : Cancel() seul ne suffit pas, le ReadString
        /// NI-VISA bloque jusqu'à la fin de la gate (jusqu'à 1000 s), le SDC le fait rendre la main en quelques ms.
        /// </summary>
        public void AborterMesureEnCours() => _driver.AborterToutesSessions();

        public async Task<ResultatMesure> ExecuterAsync(
            Mesure mesure,
            Rubidium rubidium,
            double? fNominaleOuReference,
            IProgress<ProgressionMesure>? progress = null,
            CancellationToken ct = default)
        {
            var result = new ResultatMesure();

            // Profiling accumulé en mémoire (Stopwatch + List.Add, ~10 ns par marqueur).
            // Avant, les 30+ Perf() par gate écrivaient en SQL en live : 5-10 ms l'écriture,
            // jusqu'à 300 ms d'overhead par gate sur les mesures rapides.
            var profiler = new ProfilerSession();
            void Perf(string label) => profiler.Mark(label);

            // déclarées avant le try : le nettoyage d'annulation (catch/finally) en a besoin
            string? derniereFeuille = null;
            // vrai entre la création de la feuille d'une gate et la fin de son acquisition :
            // en cas d'arrêt utilisateur, on sait si la feuille courante est incomplète (à
            // supprimer) ou si c'est une gate déjà finie (à garder)
            bool feuilleCouranteIncomplete = false;
            // chemin du .xlsx, visible dans le finally pour la suppression sur disque après un stop
            string? cheminFichier = null;
            // valeur hors domaine du module d'incertitude : on supprime la feuille courante
            // comme pour un arrêt utilisateur
            bool horsModule = false;

            try
            {
                Perf("START");

                // 0. Reset des sessions GPIB cachées par le driver VISA : évite les timeouts
                //    quand un appareil a été éteint/rallumé entre deux mesures.
                _driver.ReinitialiserSessions();
                Perf("ReinitialiserSessions");

                // 1. Résolution de l'appareil via le catalogue local.
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

                // 2. Initialisation de l'appareil + config (IFC, ConfEntree, SRQ on, rejeu SCPI)
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
                    // init manuelle : pas de *RST/*CLS, ça remettrait l'appareil aux défauts
                    // et écraserait la config faite à la main par l'opérateur
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

                // 3. Boucle de balayage : 1 itération pour la plupart des mesures, N pour une
                //    stabilité multi-gate. Chaque itération crée sa propre feuille Excel et
                //    sa propre ligne dans Récap.
                var gates = mesure.GateIndices.Count > 0
                    ? mesure.GateIndices.ToList()
                    : new List<int> { 6 };  // Sécurité : 1 s par défaut si la liste est vide.

                int nbIterations = gates.Count;
                int totalEtapes = nbIterations * (mesure.NbMesures + 3);
                int etape = 0;

                bool bulkDejaEchoue = false;

                // Mode Excel invisible (historique Stabilité) : tout en mémoire via ClosedXML,
                // le classeur n'est ouvert dans Excel qu'à la fin du balayage. Gain : ~150 ms
                // d'open Interop + ~80 ms de close par gate, soit 4-5 s sur 8 gates.
                // OPTION A : désormais tous les types (Freq + Stab) passent en mode visible avec
                // finalisation 100% COM, plus de fenêtre Frame grise en fin de mesure. Le graphe
                // Stab est patché via le Chart Object Model COM, le chemin invisible ne sert plus.
                bool excelInvisible = false;
                if (mesure.TypeMesure == TypeMesure.Stabilite)
                {
                    // Reset des sigmas accumulés (calibration axe Y graphe Stab à la dernière gate).
                    ExcelInteropHost.Instance.ReinitialiserSigmasStab();
                }

                for (int g = 0; g < nbIterations; g++)
                {
                    // Stop utilisateur entre deux gates : on ne démarre pas la suivante.
                    // Sans ce check, un balayage stab 8 gates peut continuer à enchaîner les
                    // gates restantes même après Cancel() si le ReadString de la gate courante
                    // a déjà rendu la main avant le Device Clear.
                    ct.ThrowIfCancellationRequested();

                    bool isDernier = g == nbIterations - 1;
                    int gateIdx = gates[g];
                    Perf($"--- Gate {g + 1}/{nbIterations} ({EnTetesMesureHelper.LibelleGate(gateIdx)}) ---");

                    // --- Décision : voie COM pure (option A v2) ou voie ClosedXML (chemin historique) ---
                    // En mode visible, si Excel a déjà un classeur ouvert pour CETTE FI, on ajoute la
                    // nouvelle feuille via COM pur au lieu de fermer/rouvrir (la fermeture provoque la
                    // shell SDI grise visible 1-3 s entre 2 mesures). Le Workbook ne se ferme jamais,
                    // donc plus de moment 0-Workbook ni de shell parasite.
                    // Conditions pour la voie COM :
                    //   1. mode visible (excelInvisible=false) : la Stabilité a son chemin ClosedXML
                    //      en mémoire qui n'a pas ce problème ;
                    //   2. un classeur est déjà ouvert dans Excel COM (mesure 2+ ou multi-gates) ;
                    //   3. ce classeur correspond à la mesure en cours :
                    //      - Stab : la 1ère gate (g=0) peut générer un fichier _v2 si une session
                    //        précédente existait, donc ClosedXML obligatoire ; les gates suivantes
                    //        restent dans le même classeur, COM OK.
                    //      - autres types : on compare le chemin du classeur actif au chemin attendu
                    //        pour cette FI (rejet si l'utilisateur a changé de FI).
                    bool memeMesureEnCours = false;
                    if (ExcelInteropHost.Instance.AClasseurActif)
                    {
                        if (mesure.TypeMesure == TypeMesure.Stabilite)
                        {
                            memeMesureEnCours = (g > 0);
                        }
                        else
                        {
                            string cheminAttendu = _excel.CalculerCheminFichierAttendu(mesure);
                            string cheminActif = ExcelInteropHost.Instance.CheminClasseurActif;
                            // Path.GetFullPath lève ArgumentException si l'un est vide, d'où le
                            // guard (cas 1ère mesure de la session : AClasseurActif peut être true
                            // sans que _cheminClasseurActif soit encore alimenté).
                            if (!string.IsNullOrEmpty(cheminActif) && !string.IsNullOrEmpty(cheminAttendu))
                            {
                                try
                                {
                                    memeMesureEnCours = string.Equals(
                                        Path.GetFullPath(cheminActif),
                                        Path.GetFullPath(cheminAttendu),
                                        StringComparison.OrdinalIgnoreCase);
                                }
                                catch { /* path invalide → on retombe sur voie ClosedXML */ }
                            }
                        }
                    }
                    // Voie COM v2 DÉSACTIVÉE : retour à la voie ClosedXML pour toutes les mesures,
                    // plus robuste (pas de 0x800A03EC sur les formules métier, pas de strikethrough
                    // conditionnel visible, perf équivalente). Le moment 0-Workbook-visible (cause
                    // de la shell SDI grise) est protégé par _excel.Visible=false dans
                    // FermerClasseurActifInterne + Visible=true en fin d'OuvrirEtAfficherAsync,
                    // comme à l'ouverture initiale (mesure 1) qui marche parfaitement.
                    bool voieCom = false;

                    // log de diagnostic : permet de vérifier dans le journal Admin (mode debug)
                    // quel chemin (COM ou ClosedXML) a été pris pour la mesure 2+ sur la même FI
                    JournalLog.Info(CategorieLog.Excel, "BRANCHEMENT_OPTION_A_V2",
                        $"Décision branchement : voieCom={voieCom} "
                        + $"(excelInvisible={excelInvisible}, "
                        + $"AClasseurActif={ExcelInteropHost.Instance.AClasseurActif}, "
                        + $"memeMesureEnCours={memeMesureEnCours}, "
                        + $"cheminActif='{ExcelInteropHost.Instance.CheminClasseurActif}', "
                        + $"typeMesure={mesure.TypeMesure}, gate={g})");

                    progress?.Report(new ProgressionMesure
                    {
                        Message = nbIterations > 1
                            ? $"Gate {g + 1}/{nbIterations} ({EnTetesMesureHelper.LibelleGate(gateIdx)}) — préparation Excel..."
                            : $"Préparation du rapport Excel pour FI {mesure.NumFI}...",
                        EtapeActuelle = ++etape,
                        EtapesTotales = totalEtapes
                    });

                    string nomFeuille;
                    if (voieCom)
                    {
                        // option A v2 : voie COM pure (Workbook reste ouvert, pas de shell grise)
                        nomFeuille = await ExcelInteropHost.Instance.AjouterFeuilleMesureAsync(
                            mesure, rubidium, gateIdx);
                        Perf("Interop.AjouterFeuilleMesureAsync (COM pur, Workbook gardé ouvert)");
                        cheminFichier = ExcelInteropHost.Instance.CheminClasseurActif;
                        derniereFeuille = nomFeuille;
                    }
                    else
                    {
                        // Chemin historique : ClosedXML init + Sauvegarder + Ouvrir Excel
                        if (!excelInvisible)
                        {
                            await ExcelInteropHost.Instance.FermerClasseurActifAsync();
                            Perf("Interop.FermerClasseurActif (avant init ClosedXML)");
                            // Excel COM relâche son handle fichier de façon asynchrone : sans cette
                            // attente, FichierEstVerrouille dans InitialiserRapportAsync retourne
                            // true (cache MRU Excel) et on part en fallback timestamp.
                            await Task.Delay(500);
                        }

                        // 3.a Création d'une nouvelle feuille de mesure (Stab1, Stab2, … selon le slot dispo).
                        // À la 1ère gate d'une session Stab, on signale "nouvelle session" pour qu'un
                        // suffixe _v2, _v3… soit appliqué si le fichier précédent existe déjà : évite
                        // que le graphe Stab traîne les 7 valeurs de la mesure précédente sur le même FI.
                        bool nouvelleSession = (g == 0 && mesure.TypeMesure == TypeMesure.Stabilite);
                        await _excel.InitialiserRapportAsync(mesure.NumFI, mesure, rubidium, gateIdx, nouvelleSession);
                        Perf("ClosedXML.InitialiserRapportAsync");
                        await _excel.PreparerLignesMesureAsync(mesure.NbMesures);
                        Perf("ClosedXML.PreparerLignesMesureAsync");

                        cheminFichier = await _excel.SauvegarderSurDisqueAsync();
                        Perf("ClosedXML.SauvegarderSurDisqueAsync (.xlsx)");
                        nomFeuille = _excel.NomFeuilleMesure;
                        derniereFeuille = nomFeuille;

                        // En mode visible : on ferme ClosedXML et on ouvre Excel pour l'utilisateur.
                        // En mode invisible (Stabilité) : ClosedXML reste prêt en mémoire, on
                        // remplit le tableau de mesures dedans, sans toucher à Excel.
                        if (!excelInvisible)
                        {
                            _excel.FermerExcel();
                            Perf("ClosedXML.FermerExcel");
                            await ExcelInteropHost.Instance.OuvrirEtAfficherAsync(cheminFichier, nomFeuille);
                            Perf("Interop.OuvrirEtAfficherAsync");
                        }
                    }

                    // La feuille de cette gate vient d'être créée mais ses valeurs ne sont pas
                    // encore acquises : si l'utilisateur stoppe maintenant, il faudra la supprimer
                    // (cf. catch OperationCanceledException). Repassé à false en fin de gate.
                    feuilleCouranteIncomplete = true;

                    // 3.b Programmation de la gate de cette itération
                    progress?.Report(new ProgressionMesure
                    {
                        Message = nbIterations > 1
                            ? $"Gate {g + 1}/{nbIterations} ({EnTetesMesureHelper.LibelleGate(gateIdx)}) — programmation..."
                            : $"Programmation de la gate (index {gateIdx})...",
                        EtapeActuelle = ++etape,
                        EtapesTotales = totalEtapes
                    });

                    // Entre deux gates d'un balayage : on aborte la mesure en cours et clear
                    // le statut pour repartir d'un buffer d'output vide. Sans ça, le 53131A
                    // (et la plupart des compteurs) gardent l'état FETCH de la gate précédente
                    // et le :READ:FREQ? suivant timeout (l'instrument croit qu'il a déjà répondu).
                    // Diagnostic empirique : la 1ère gate (post-*RST) marche, les suivantes
                    // timeout systématiquement sans cette séquence de réinit légère.
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

                    // Skip la vérification d'arming dans les balayages stab (déjà éprouvée par
                    // les itérations précédentes, ~200 ms de gain par gate). Reste activée pour
                    // les autres types de mesure et pour la 1ère gate d'un balayage.
                    bool verifierArming = !(mesure.TypeMesure == TypeMesure.Stabilite && g > 0);
                    await appareil.AppliquerGateAsync(_driver, gateIdx, mesure.TypeMesure, ct, verifierArming);
                    Perf($"AppliquerGateAsync (verifArming={verifierArming})");

                    // 3.c Boucle de N mesures.
                    //   - Stabilité avec gate COURTE (< 500 ms) : on accumule en RAM pour écrire
                    //     en bloc à la fin de la gate. À 10 ms de gate la boucle dure ~1 s, l'œil
                    //     ne suit pas l'affichage cellule par cellule, et chaque écriture COM live
                    //     coûte ~300 ms (re-render Excel visible). Gain typique : 22 s → 3-4 s.
                    //   - Stabilité avec gate LONGUE (≥ 500 ms) : écriture live, la mesure dure
                    //     plusieurs secondes de toute façon donc l'overhead (~300 ms par mesure)
                    //     est marginal, autant que l'utilisateur voie les valeurs arriver.
                    //   - Autres types (Fréquence, Interval…) : toujours en live, cadence humaine.
                    double gateSecondes = EnTetesMesureHelper.SecondesGate(gateIdx);
                    bool ecritureBatch = mesure.TypeMesure == TypeMesure.Stabilite
                                         && gateSecondes < 0.5;
                    var valeurs = new List<double>();
                    var bufferBatch = ecritureBatch ? new List<(DateTime, double)>(mesure.NbMesures) : null;
                    const int LIGNE_DEBUT_MESURES = 9;

                    // Stratégie d'acquisition (3 niveaux, du plus rapide au plus standard) :
                    //   1. Bulk    : CommandeMesureMultiple (avec {N}), l'instrument fait
                    //                les N mesures en interne et les retourne en CSV.
                    //                ~0,5 s pour 30 mesures à 10 ms (vs ~6 s en :FETCh?).
                    //   2. Rapide  : :INIT:CONT ON + boucle :FETCh?, évite le ré-arming
                    //                de :READ? (~180 ms/mesure vs ~670 ms).
                    //   3. Classic : :READ? à chaque mesure (ré-arme à chaque appel).
                    // Le choix se fait depuis le catalogue (champ CommandeMesureMultiple),
                    // aucune logique spécifique à un modèle d'appareil dans ce code.
                    string? cmdFetch = AppareilIeeeOperations.DeriverCommandeFetch(appareil.ExeMesure);

                    // Si l'utilisateur a renseigné une commande "fetch fresh" (qui bloque jusqu'à
                    // nouvelle mesure dispo), on l'utilise prioritairement et on supprime tout
                    // Task.Delay : c'est l'instrument qui pace, pas nous.
                    bool fetchBloquant = !string.IsNullOrWhiteSpace(appareil.CommandeFetchFresh);
                    string? cmdFetchUsed = fetchBloquant ? appareil.CommandeFetchFresh : cmdFetch;

                    // Toute mesure avec NbMesures > 1 = série exploitable par streaming/mode
                    // rapide. Cas spécial NbMesures == 1 : streaming/rapide n'apportent rien
                    // (overhead bulk init pour 1 mesure), on garde le path :READ? classique.
                    //
                    // Les sécurités cumulatives en dessous (CommandeBulkInit/FetchFresh non
                    // vides côté catalogue, gateSecondes ≥ 1 s, bulkDejaEchoue) font de toute
                    // façon le tri : un instrument qui ne supporte pas le streaming pour un
                    // type donné n'engage pas le mode, on retombe sur le fallback classique.
                    bool typeFaitSeries = mesure.NbMesures > 1;

                    // Mode streaming : acquisition gap-free + lecture une-par-une via
                    // CommandeFetchFresh (typ. ":DATA:REM? 1,WAIT" sur 53230A). Compatible
                    // avec live UI ET Stab batch. Réservé aux gates ≥ 1 s : sous ce seuil,
                    // chaque DATA:REM coûte ~30-50 ms de roundtrip GPIB qui polluerait la
                    // cadence de mesures rapides, on préfère le bulk classique.
                    bool modeStreamingPossible = !string.IsNullOrWhiteSpace(appareil.CommandeBulkInit)
                        && !string.IsNullOrWhiteSpace(appareil.CommandeFetchFresh)
                        && gateSecondes >= 1.0
                        && typeFaitSeries
                        && !bulkDejaEchoue;

                    // Mode bulk : timestamps simulés (l'instrument fait les mesures en interne sans
                    // nous prévenir), incompatible avec l'écriture live Excel cellule par cellule.
                    // On le réserve donc à la Stabilité (où on bufferise de toute façon).
                    // Désactivé si le streaming est utilisé (les 2 sont mutuellement exclusifs).
                    bool modeBulk = !modeStreamingPossible
                        && !string.IsNullOrWhiteSpace(appareil.CommandeMesureMultiple)
                        && mesure.TypeMesure == TypeMesure.Stabilite
                        && !bulkDejaEchoue;
                    // Mode rapide :FETCh? : applicable à tout type avec NbMesures > 1 dès lors
                    // que l'appareil l'autorise via ModeRapideActif du catalogue. Compatible
                    // avec l'écriture live Excel. Gain ~3× vs :READ? classique sur le 53131A.
                    bool modeRapide = !modeBulk
                        && !modeStreamingPossible
                        && !string.IsNullOrEmpty(cmdFetchUsed)
                        && typeFaitSeries
                        && appareil.ModeRapideActif;   // bloqué pour 53230A & co (cf. catalogue)

                    // Délai inter-fetch :
                    //   - 0 si la commande est bloquante (l'instrument fait l'attente)
                    //   - Sinon : gate - 100 ms (plancher 5 ms). Le write+read GPIB prend
                    //     ~150 ms, donc Delay+150 ≥ gate garantit une nouvelle mesure dispo
                    //     au moment du :FETCh?. Pour gates ≤ 100 ms le plancher 5 ms suffit
                    //     (write+read masque la gate). Pour gates ≥ 200 ms on attend juste
                    //     ce qu'il faut, sans gaspiller comme ferait un gate × 0.5 trop court.
                    int delayFetchMs;
                    if (!modeRapide) delayFetchMs = 0;
                    else if (fetchBloquant) delayFetchMs = 0;
                    else delayFetchMs = Math.Max(5, (int)(gateSecondes * 1000) - 100);

                    if (modeRapide)
                    {
                        try
                        {
                            await _driver.EcrireAsync(appareil.Adresse, ":INIT:CONT ON", appareil.WriteTerm, ct);
                            // Amorçage adaptatif : délai initial = gate + 200 ms (mini 150 ms),
                            // garantit que la 1ère mesure interne est complète avant le 1er fetch.
                            // Si le fetch retourne 0 (erreur GPIB 230 "Data stale"), on retry en
                            // augmentant le délai (2x, 3x, 4x).
                            int amorceInitial = Math.Max(150, (int)(gateSecondes * 1000) + 200);
                            await Task.Delay(amorceInitial, ct);
                            for (int essai = 0; essai < 3; essai++)
                            {
                                var testVal = await appareil.FetcherAsync(_driver, cmdFetchUsed!, ct);
                                if (testVal != 0.0) break;
                                await Task.Delay(amorceInitial * (essai + 2), ct); // 2x, 3x, 4x
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

                    // Déclaré ici pour rester accessible après le goto FinBoucleMesures
                    // (le bloc bulk ci-dessous peut sauter directement à la fin sans entrer
                    // dans la boucle classique).
                    int totalDoublons = 0;

                    // Écritures Excel via Interop COM lancées en fire-and-forget : chaque
                    // EcrireValeurLiveAsync prend ~50-100 ms (lock COM), ce qui ajoute autant
                    // de latence par mesure si on les await en série. En parallélisant avec
                    // la mesure GPIB suivante (gate ~1 s), on récupère ce temps et on retombe
                    // sur la cadence physique de l'instrument. Le lock(_sync) interne
                    // d'ExcelInteropHost garantit la sérialisation FIFO des écritures (à
                    // cadence ≤ 10 Hz typique du fréquencemètre, pas de race possible).
                    // Vide en mode batch (Stab) — on ne touche pas à Interop pendant les gates.
                    var pendingWrites = new List<Task>(mesure.NbMesures);

                    // ===== MODE STREAMING : prioritaire sur tous les autres si applicable =====
                    // Acquisition gap-free + lecture une-par-une via CommandeFetchFresh
                    // (typ. ":DATA:REM? 1,WAIT" sur 53230A). Compatible avec live UI ET batch
                    // Stab. Réservé aux gates ≥ 1 s (en dessous, le DATA:REM coûte ~30-50 ms
                    // de roundtrip GPIB qui polluerait la cadence).
                    if (modeStreamingPossible)
                    {
                        // Timeout : gate × N + 3 s, comme le bulk classique.
                        int timeoutStream = Math.Max(5000,
                            (int)(gateSecondes * 1000 * mesure.NbMesures) + 3000);
                        _driver.DefinirTimeout(appareil.Adresse, timeoutStream);

                        var swStream = Stopwatch.StartNew();
                        bool streamOk = false;
                        try
                        {
                            // 1. Lancer l'acquisition gap-free en arrière-plan.
                            string cmdInit = appareil.CommandeBulkInit.Replace("{N}",
                                mesure.NbMesures.ToString(CultureInfo.InvariantCulture));
                            await _driver.EcrireAsync(appareil.Adresse, cmdInit, appareil.WriteTerm, ct);

                            // 2. Boucle : lire chaque valeur dès qu'elle est dispo (WAIT bloque
                            //    jusqu'à dispo dans la mémoire de readings du compteur).
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

                                // Throttle UI uniforme : 1/5 + dernière mesure (cf. boucle classique).
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
                            streamOk = true;
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
                        // Timeout adapté au volume réel : gate × N + 3 s de marge (init mode
                        // CONT, transmission CSV des N valeurs, parsing).
                        int timeoutBulk = Math.Max(5000,
                            (int)(gateSecondes * 1000 * mesure.NbMesures) + 3000);
                        _driver.DefinirTimeout(appareil.Adresse, timeoutBulk);

                        // Message UI clair : en mode bulk, l'instrument fait les N mesures en
                        // interne sans qu'on puisse les remonter une par une à l'utilisateur.
                        double dureeAttendueSec = gateSecondes * mesure.NbMesures;
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

                        try
                        {
                            var lot = await appareil.MesurerEnLotAsync(
                                _driver, appareil.CommandeMesureMultiple, mesure.NbMesures, ct);
                            int nbRecu = lot.Count;
                            swBulkProgress.Stop();

                            // On considère le bulk OK si on a reçu au moins 80 % des mesures
                            // attendues. Sinon (instrument refuse la commande, retour vide,
                            // erreur SCPI), on bascule sur le mode :FETCh? rapide.
                            if (nbRecu >= mesure.NbMesures * 0.8)
                            {
                                JournalLog.Info(CategorieLog.Mesure, "BULK_FETCH",
                                    $"Lot reçu : {nbRecu}/{mesure.NbMesures} valeurs en 1 transaction GPIB.");

                                // Message UI : confirme à l'utilisateur que les valeurs ont
                                // bien été acquises et combien de temps ça a pris.
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

                        // Mémorise l'échec : on n'essaiera plus le bulk pour les gates suivantes.
                        bulkDejaEchoue = true;
                        JournalLog.Info(CategorieLog.Mesure, "BULK_DESACTIVE_POUR_BALAYAGE",
                            "Commande bulk désactivée pour le reste du balayage — la 1ère tentative a échoué.");

                        // Fallback léger : juste :ABORT + *CLS pour purger une éventuelle query
                        // interrompue côté instrument. Pas de *RST complet (l'instrument n'est
                        // pas corrompu, juste la commande inconnue) → on garde les réglages
                        // SCPI utilisateur en place, pas besoin de tout rejouer.
                        try
                        {
                            await _driver.EcrireAsync(appareil.Adresse, ":ABORT", appareil.WriteTerm, ct);
                            await Task.Delay(20, ct);
                            await _driver.EcrireAsync(appareil.Adresse, "*CLS", appareil.WriteTerm, ct);
                            await Task.Delay(50, ct);
                        }
                        catch { /* best effort */ }

                        // Restaure le timeout VISA normal pour les opérations suivantes
                        // (basé sur la gate, comme dans AppliquerGateAsync).
                        int timeoutNormal = Math.Max(5000, (int)(gateSecondes * 1000) + 2000);
                        _driver.DefinirTimeout(appareil.Adresse, timeoutNormal);

                        modeBulk = false;
                        modeRapide = !string.IsNullOrEmpty(cmdFetch);
                        if (modeRapide)
                        {
                            // Recalcule delayFetchMs (était à 0 car modeRapide était false avant le fallback).
                            delayFetchMs = Math.Max(15, (int)(gateSecondes * 1000) + 5);
                            try
                            {
                                await _driver.EcrireAsync(appareil.Adresse, ":INIT:CONT ON", appareil.WriteTerm, ct);
                                await Task.Delay(delayFetchMs + 50, ct);
                            }
                            catch { modeRapide = false; }
                        }
                        // La boucle classique ci-dessous va maintenant utiliser modeRapide ou :READ?
                    }

                    // Détection des doublons : si :FETCh? retourne la même valeur que la
                    // précédente, on attend gate/2 supplémentaire et on retry. Max 3 retries
                    // pour ne pas boucler infiniment si l'instrument est figé.
                    //
                    // Détection des timeouts FetchFresh : si la commande "fresh" retourne 0
                    // (= timeout VISA + Device Clear) 3 fois d'affilée, c'est qu'elle n'est
                    // pas reconnue par l'instrument — on désactive et on bascule en :FETCh?
                    // classique avec Task.Delay (sans avoir à arrêter la mesure).
                    double valPrec = double.NaN;
                    int doublonsConsecutifs = 0;
                    int freshTimeoutsConsec = 0;
                    for (int i = 0; i < mesure.NbMesures; i++)
                    {
                        ct.ThrowIfCancellationRequested();

                        double val;
                        if (modeRapide)
                        {
                            if (i > 0 && delayFetchMs > 0) await Task.Delay(delayFetchMs, ct);
                            val = await appareil.FetcherAsync(_driver, cmdFetchUsed!, ct);

                            // Auto-fallback FetchFresh : si la commande bloquante timeout
                            // 3 fois d'affilée, on bascule sur :FETCh? classique + delay.
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
                                    // Re-arme le mode continu (au cas où FetchFresh l'aurait perturbé)
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

                            // DÉTECTION DOUBLON RETIRÉE (cause des sauts de 3-4 s observés
                            // sur signal stable, ex: générateur carré propre où chaque mesure
                            // donne le même résultat au bit près). L'heuristique « si même
                            // valeur que la précédente, c'est que l'instrument est figé »
                            // pénalise injustement les sources stables, qui sont précisément
                            // ce qu'on cherche à mesurer en métrologie.
                            //
                            // Le cas légitime (FETCh? appelé trop vite, relit ancienne valeur)
                            // est déjà couvert en amont par `delayFetchMs = gate - 100 ms` qui
                            // garantit qu'une nouvelle gate est complète avant la prochaine
                            // lecture. Sur gate ≥ 1 s la marge est de 900 ms, donc aucun risque
                            // de relire l'ancienne valeur. Sur gates très courtes (10/100 ms),
                            // l'instrument lui-même rafraîchit ses registres à chaque cycle.
                            //
                            // Comportement avant retrait : 3 retries × gate/2 = +1,5 s par mesure
                            // stable détectée → sur 30 mesures, jusqu'à +45 s ajoutées au temps
                            // total. Maintenant : les valeurs identiques sont conservées telles
                            // quelles, c'est de la donnée valide pas un bug.
                            valPrec = val;
                            // doublonsConsecutifs / totalDoublons gardés pour compat journal
                            // mais ne s'incrémentent plus.
                        }
                        else
                        {
                            // Borne l'attente du MAV sur la gate (+1 s de marge) : assez pour qu'un
                            // appareil dont le MAV fonctionne le positionne (il l'assert ~fin de gate),
                            // mais sans plancher à 5 s. Si le MAV ne vient pas, MesurerAsync le détecte
                            // et bascule en lecture directe pour les points suivants (cadence = gate).
                            int mavTimeoutMs = Math.Max(1500, (int)(gateSecondes * 1000) + 1000);
                            val = await appareil.MesurerAsync(_driver, mesure, ct, mavTimeoutMs);
                        }
                        var ts = DateTime.Now;
                        valeurs.Add(val);

                        if (ecritureBatch)
                        {
                            bufferBatch!.Add((ts, val));
                        }
                        else
                        {
                            // Fire-and-forget : l'écriture Excel via Interop tourne en
                            // parallèle de la prochaine mesure GPIB. Flushé en fin de boucle
                            // par Task.WhenAll(pendingWrites).
                            pendingWrites.Add(ExcelInteropHost.Instance.EcrireValeurLiveAsync(i, val, ts));
                        }

                        // Throttle UNIFORME (Stab + Freq) : 1 sur 5 + dernière mesure. Chaque
                        // Report() trigge un re-render du TextBox UI sur le thread dispatcher,
                        // ce qui peut bloquer l'UI 200-500 ms quand InformationsGenerales fait
                        // plusieurs Ko. Sans throttle uniforme, en mode Freq les 30 reports/gate
                        // saturaient le UI thread et le bouton « Arrêter » devenait incliquable
                        // pendant la mesure (le click était mis en queue jusqu'à libération).
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

                    // Flush des écritures Excel encore en vol — la boucle GPIB peut finir
                    // avant que les dernières cellules soient écrites côté Interop COM.
                    // Sans ce WhenAll, on enchaînerait sur EcrireStatsAsync alors qu'Excel
                    // n'a pas encore reçu toutes les valeurs → cellules vides côté user.
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
                    string suffixe = modeRapide && totalDoublons > 0 ? $" / {totalDoublons} doublon(s) re-essayé(s)" : "";
                    Perf($"Boucle GPIB {mesure.NbMesures} mesures ({swBoucle.ElapsedMilliseconds} ms total, ~{swBoucle.ElapsedMilliseconds / Math.Max(1, mesure.NbMesures)} ms/mesure — {modeTag}{suffixe})");

                    // Désactive le mode continu après la boucle pour ne pas laisser
                    // l'instrument trigger en arrière-plan entre les gates ou après la mesure.
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

                    // Flush batch : écriture en bloc des N valeurs accumulées.
                    //   - Mode invisible (Stabilité) : via ClosedXML directement en mémoire,
                    //     pas d'aller-retour Excel/COM.
                    //   - Mode visible (Fréquence batch ou autre) : via Interop (le classeur
                    //     est ouvert dans Excel, on bénéficie du Range.Value2 = matrix).
                    if (ecritureBatch && bufferBatch!.Count > 0)
                    {
                        if (excelInvisible)
                        {
                            await _excel.EcrireValeursBatchClosedXMLAsync(LIGNE_DEBUT_MESURES, bufferBatch);
                            Perf("ClosedXML.EcrireValeursBatch (mode invisible)");
                        }
                        else
                        {
                            await ExcelInteropHost.Instance
                                .EcrireValeursEnBlocAsync(LIGNE_DEBUT_MESURES, bufferBatch);
                            Perf("Interop.EcrireValeursEnBlocAsync (bulk write)");
                        }
                    }

                    // 3.d SRQ off uniquement à la dernière itération (cf. F_Main:1262 — historiquement
                    //     un correctif Racal, mais le principe reste : on évite de désactiver puis
                    //     réactiver SRQ entre deux gates consécutives).
                    if (isDernier)
                    {
                        await appareil.DesactiverSrqAsync(_driver, ct);
                        Perf("DesactiverSrqAsync");
                    }

                    // 3.e Stats + ligne Recap pour cette gate
                    //   - Mode visible (Freq/Interval/Tachy) : OPTION A — tout via COM directement
                    //     dans le Workbook ouvert. Pas de cycle ferme/rouvre → plus de fenêtre Frame
                    //     grise à la fin de la mesure.
                    //   - Mode invisible (Stab) : chemin historique ClosedXML (graphe Stab nécessite
                    //     l'accès ClosedXML pour ses caches numCache).
                    progress?.Report(new ProgressionMesure
                    {
                        Message = nbIterations > 1
                            ? $"Gate {g + 1}/{nbIterations} — finalisation..."
                            : "Finalisation du rapport Excel...",
                        EtapeActuelle = etape,
                        EtapesTotales = totalEtapes
                    });

                    // Garde-fou « dépassement de module » : si la moyenne de cette gate sort du
                    // domaine couvert par le module d'incertitude sélectionné, on n'écrit pas la
                    // feuille — on lève une exception typée. Le finally supprimera la feuille en
                    // cours (comme un arrêt) et le ViewModel affichera la popup d'erreur.
                    // (Module introuvable ≠ hors plage : on laisse passer le fallback historique.)
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

                    if (!excelInvisible)
                    {
                        // OPTION A : tout en COM, pas de fermeture/réouverture.
                        // Pour Stab : à la dernière gate, le graphe est patché aussi via COM
                        // (plages séries + bornes axe Y log) → plus besoin du cycle XML ClosedXML.
                        int nbGatesStab = mesure.TypeMesure == TypeMesure.Stabilite ? (g + 1) : 0;
                        await ExcelInteropHost.Instance.FinaliserMesureViaComAsync(
                            mesure, valeurs, nomFeuille, gateSecondes, isDernier, nbGatesStab);
                        Perf("Interop.FinaliserMesureViaComAsync");

                        // PLUS de duplication fichier-par-fichier ici : le transfert vers le
                        // réseau est fait en BLOC à la toute fin de la mesure via
                        // TransfertReseauService.TransfererDossierFIAsync — qui copie le dossier
                        // FI complet (xlsm + Journal_FI.txt + profilings + tout).
                    }
                    else
                    {
                        // Mode invisible (Stab) : chemin ClosedXML historique
                        await _excel.EcrireStatsAsync(valeurs);
                        Perf("ClosedXML.EcrireStatsAsync");

                        if (mesure.TypeMesure == TypeMesure.Stabilite)
                        {
                            await _excel.MettreAJourRecapStabAsync(mesure);
                            Perf("ClosedXML.MettreAJourRecapStabAsync");
                        }

                        await _excel.SauvegarderFinalAsync();
                        Perf("ClosedXML.SauvegarderFinalAsync");

                        if (isDernier)
                        {
                            _excel.FermerExcel();
                            Perf("ClosedXML.FermerExcel (final)");
                            await ExcelInteropHost.Instance.OuvrirEtAfficherAsync(cheminFichier, nomFeuille);
                            Perf("Interop.OuvrirEtAfficherAsync (1 seule fois à la fin du balayage)");
                        }
                    }

                    // À la dernière itération d'un balayage de stabilité, on ajoute le graphe
                    // de stabilité dans la Récap. (3 séries : Écart type / Valeurs Maxi / Mini
                    // en fonction du Temps de porte). Idempotent côté Interop, et on save
                    // pour persister le graphe dans le fichier final.
                    // Le graphe de stabilité est désormais embarqué directement dans le template
                    // METROLOGO_Stab.xlsx (extrait de Stab1.xls historique avec toutes ses
                    // propriétés exactes : HiLoLines, log Y, markers Diamond/Dash/Dot, légende
                    // en bas, etc.). Plus besoin de le recréer programmatiquement — il est cloné
                    // automatiquement à chaque nouveau classeur de mesure stabilité.

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

                    // Gate terminée : sa feuille est complète. Si un arrêt survient pendant la
                    // gate suivante, on ne doit PAS supprimer cette feuille-ci.
                    feuilleCouranteIncomplete = false;
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

                // Transfert du dossier FI complet vers le partage réseau M:\ (snapshot en bloc).
                // Inclut tout : Mesures_FI.xlsx + Mesures_Stab_FI.xlsx + Journal_FI.txt + profilings.
                // Si KO (M:\ down, latence, etc.) : FI inscrite dans la liste de reprise pour
                // rejouer au prochain démarrage. Le ViewModel affiche un MessageBox à l'utilisateur.
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
                // Valeur hors du domaine du module : feuille à supprimer (cf. finally), popup
                // côté ViewModel via result.Erreur. result.Succes reste false.
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
                // Mesure stoppée : la feuille créée pour la mesure en cours est incomplète et ne
                // doit pas être conservée. On déclenche la suppression dès que l'annulation a été
                // demandée — PEU IMPORTE l'exception remontée : un arrêt via Device Clear interrompt
                // un :READ? bloquant en levant une exception VISA (et non OperationCanceledException),
                // qui partait jusqu'ici dans le catch générique sans nettoyer la feuille.
                // On ne supprime QUE si la feuille courante est incomplète (sinon, en balayage
                // stabilité, on effacerait une gate précédente déjà terminée).
                if ((ct.IsCancellationRequested || horsModule)
                    && feuilleCouranteIncomplete
                    && !string.IsNullOrEmpty(derniereFeuille))
                {
                    // 1. Tentative COM (utile si Excel est encore vivant — ex. arrêt sans kill).
                    try { await ExcelInteropHost.Instance.SupprimerFeuilleMesureAsync(derniereFeuille); }
                    catch (Exception exSuppr)
                    {
                        JournalLog.Warn(CategorieLog.Mesure, "Execute_SupprFeuilleStop",
                            $"Suppression COM de la feuille « {derniereFeuille} » après arrêt échouée : {exSuppr.Message}");
                    }

                    // 2. Suppression SUR DISQUE via ClosedXML : le bouton STOP tue le process Excel
                    //    (TuerProcessExcelAsync), donc la suppression COM ci-dessus est un no-op. Le
                    //    fichier .xlsx sur disque contient déjà la feuille sauvegardée → on le nettoie
                    //    directement, sans Excel. C'est le chemin réellement efficace après un STOP.
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

                // NOTE : l'écriture du fichier Profiling_<FI>_<date>.txt dans le dossier de la FI
                // a été DÉSACTIVÉE — il polluait les dossiers de mesure et n'est plus consulté.
                // Le profiler reste alimenté en mémoire (appels Perf(...)) au cas où on voudrait
                // un jour le ré-exposer (ex. log central), mais plus aucun fichier n'est créé à
                // côté des rapports Excel.
            }

            return result;
        }

        /// <summary>
        /// Rejoue les commandes SCPI des réglages dynamiques choisis dans Configuration. Sans ça
        /// le <c>*RST</c> de la ChaineInit efface tous les réglages que l'utilisateur a validés.
        /// Chaque commande est wrappée : si une timeout, on logge et on continue les suivantes
        /// (plutôt que de planter toute la mesure). Délai de 50 ms entre commandes pour les
        /// appareils lents (53131A).
        /// </summary>
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
                // Catch large : VISA lève Ivi.Visa.IOTimeoutException, NI-488.2 lève IOException
                // avec « timeout » dans le message. Dans les deux cas on log et on continue.
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

        /// <summary>
        /// Retrouve le modèle catalogue + l'adresse GPIB détectée sur le bus, et construit
        /// un <see cref="AppareilIEEE"/> via <see cref="CatalogueAdapter"/>.
        /// </summary>
        private static (AppareilIEEE?, string?) ResolverDepuisCatalogue(string idModele, Mesure mesure)
        {
            var modele = CatalogueAppareilsService.Instance.Modeles
                .FirstOrDefault(m => m.Id == idModele);
            if (modele == null)
            {
                return (null, $"Le modèle catalogue « {idModele} » est introuvable. "
                    + "Le catalogue a peut-être été modifié — rouvrez la Configuration.");
            }

            // Mode « adresses fixes » : adresse explicitement forcée sur la mesure, ou appareil
            // legacy (sans *IDN?) avec son adresse par défaut. On court-circuite alors la détection
            // sur le bus (ces appareils n'apparaissent jamais dans AppareilsDetectes).
            int adresseFixe = mesure.AdresseFixeForcee >= 0
                ? mesure.AdresseFixeForcee
                : (modele.Parametres.Legacy ? modele.Parametres.AdresseFixeParDefaut : -1);

            if (adresseFixe >= 0)
                return (CatalogueAdapter.VersAppareilIEEE(modele, adresseFixe), null);

            // Chemin historique : appareil moderne détecté par IDN sur le bus.
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
