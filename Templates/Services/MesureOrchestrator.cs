using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Metrologo.Models;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Ieee;
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

        public async Task<ResultatMesure> ExecuterAsync(
            Mesure mesure,
            Rubidium rubidium,
            double? fNominaleOuReference,
            IProgress<ProgressionMesure>? progress = null,
            CancellationToken ct = default)
        {
            var result = new ResultatMesure();

            // -------- Profiling : chronomètre chaque étape et logue la durée dans le journal --------
            // Chaque appel à Perf("...") logue le temps écoulé depuis le précédent Perf, avec un
            // total cumulé (T=). Utile pour identifier où va le temps. Filtre = catégorie Mesure
            // + action PERF dans la base sqlite.
            var swPerf = Stopwatch.StartNew();
            long lastT = 0;
            void Perf(string label)
            {
                long now = swPerf.ElapsedMilliseconds;
                JournalLog.Info(CategorieLog.Mesure, "PERF",
                    $"+{(now - lastT),5} ms  T={now,6} ms  {label}");
                lastT = now;
            }

            try
            {
                Perf("START");

                // 0. Réinitialise les sessions GPIB cachées par le driver VISA — évite les
                //    timeouts quand un appareil a été éteint/rallumé entre deux mesures.
                _driver.ReinitialiserSessions();
                Perf("ReinitialiserSessions");

                // 1. Résolution de l'appareil via le catalogue local.
                if (string.IsNullOrWhiteSpace(mesure.IdModeleCatalogue))
                {
                    result.Erreur = "Aucun appareil sélectionné pour la mesure. "
                        + "Enregistrez le fréquencemètre via Administration → Gérer les appareils.";
                    return result;
                }

                var (appareil, erreur) = ResolverDepuisCatalogue(mesure.IdModeleCatalogue);
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

                await appareil.InitialiserAsync(_driver, ct);
                Perf("InitialiserAsync (*RST;*CLS)");
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

                string? cheminFichier = null;
                string? derniereFeuille = null;

                bool bulkDejaEchoue = false;

                // Mode « Excel invisible » : pour la Stabilité, Excel ne s'ouvre PAS pendant
                // les gates. Tout est fait en mémoire via ClosedXML, et le classeur final
                // n'est ouvert dans Excel qu'à la toute fin du balayage. Économie : ~150 ms
                // d'open Interop × N gates + ~80 ms de close Interop × N gates + le coût COM
                // de chaque toggle Visible. Sur 8 gates ≈ 4-5 s gagnés.
                bool excelInvisible = mesure.TypeMesure == TypeMesure.Stabilite;

                for (int g = 0; g < nbIterations; g++)
                {
                    bool isDernier = g == nbIterations - 1;
                    int gateIdx = gates[g];
                    Perf($"--- Gate {g + 1}/{nbIterations} ({EnTetesMesureHelper.LibelleGate(gateIdx)}) ---");

                    // En mode visible : Interop tient le classeur depuis la gate précédente,
                    // on le ferme pour libérer le handle avant que ClosedXML touche au fichier.
                    // En mode invisible : aucun handle Interop, rien à fermer.
                    if (g > 0 && !excelInvisible)
                    {
                        await ExcelInteropHost.Instance.FermerClasseurActifAsync();
                        Perf("Interop.FermerClasseurActif (entre gates)");
                    }

                    progress?.Report(new ProgressionMesure
                    {
                        Message = nbIterations > 1
                            ? $"Gate {g + 1}/{nbIterations} ({EnTetesMesureHelper.LibelleGate(gateIdx)}) — préparation Excel..."
                            : $"Préparation du rapport Excel pour FI {mesure.NumFI}...",
                        EtapeActuelle = ++etape,
                        EtapesTotales = totalEtapes
                    });

                    // 3.a Création d'une nouvelle feuille de mesure (Stab1, Stab2, … selon le slot dispo)
                    await _excel.InitialiserRapportAsync(mesure.NumFI, mesure, rubidium, gateIdx);
                    Perf("ClosedXML.InitialiserRapportAsync");
                    await _excel.PreparerLignesMesureAsync(mesure.NbMesures);
                    Perf("ClosedXML.PreparerLignesMesureAsync");

                    cheminFichier = await _excel.SauvegarderSurDisqueAsync();
                    Perf("ClosedXML.SauvegarderSurDisqueAsync (.xlsm)");
                    string nomFeuille = _excel.NomFeuilleMesure;
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
                    // les itérations précédentes — ~200 ms de gain par gate). Reste activée pour
                    // les autres types de mesure et pour la 1ère gate d'un balayage.
                    bool verifierArming = !(mesure.TypeMesure == TypeMesure.Stabilite && g > 0);
                    await appareil.AppliquerGateAsync(_driver, gateIdx, mesure.TypeMesure, ct, verifierArming);
                    Perf($"AppliquerGateAsync (verifArming={verifierArming})");

                    // 3.c Boucle de N mesures.
                    //   - Stabilité : on accumule les valeurs en RAM pour les écrire en bloc à
                    //     la fin de la gate (1 seul appel COM massif). À 10 ms de gate la boucle
                    //     dure ~1 s — l'œil humain ne suit pas l'affichage cellule-par-cellule
                    //     de toute façon, et chaque écriture COM live coûte ~300 ms (re-render
                    //     Excel visible). Gain typique : 22 s → 3-4 s sur les balayages courts.
                    //   - Autres types (Fréquence, Interval…) : on garde l'écriture live, la
                    //     cadence est plus humaine (gate ≥ 1 s) et l'utilisateur surveille.
                    bool ecritureBatch = mesure.TypeMesure == TypeMesure.Stabilite;
                    var valeurs = new List<double>();
                    var bufferBatch = ecritureBatch ? new List<(DateTime, double)>(mesure.NbMesures) : null;
                    const int LIGNE_DEBUT_MESURES = 9;

                    // Stratégie d'acquisition (3 niveaux, du plus rapide au plus standard) :
                    //   1. Bulk    : CommandeMesureMultiple (avec {N}) — l'instrument fait
                    //                les N mesures en interne et les retourne en CSV.
                    //                ~0,5 s pour 30 mesures à 10 ms (vs ~6 s en :FETCh?).
                    //   2. Rapide  : :INIT:CONT ON + boucle :FETCh? — évite le ré-arming
                    //                de :READ? (~180 ms/mesure vs ~670 ms).
                    //   3. Classic : :READ? à chaque mesure (ré-arme à chaque appel).
                    // Le choix se fait depuis le catalogue (champ CommandeMesureMultiple) —
                    // aucune logique spécifique à un modèle d'appareil dans ce code.
                    double gateSecondes = EnTetesMesureHelper.SecondesGate(gateIdx);
                    string? cmdFetch = AppareilIeeeOperations.DeriverCommandeFetch(appareil.ExeMesure);

                    // Si l'utilisateur a renseigné une commande "fetch fresh" (qui bloque jusqu'à
                    // nouvelle mesure dispo), on l'utilise prioritairement et on supprime tout
                    // Task.Delay — c'est l'instrument qui pace, pas nous.
                    bool fetchBloquant = !string.IsNullOrWhiteSpace(appareil.CommandeFetchFresh);
                    string? cmdFetchUsed = fetchBloquant ? appareil.CommandeFetchFresh : cmdFetch;

                    // Mode bulk : timestamps simulés (l'instrument fait les mesures en interne sans
                    // nous prévenir) — incompatible avec l'écriture live Excel cellule-par-cellule.
                    // On le réserve donc à la Stabilité (où on bufferise de toute façon).
                    bool modeBulk = !string.IsNullOrWhiteSpace(appareil.CommandeMesureMultiple)
                        && mesure.TypeMesure == TypeMesure.Stabilite
                        && !bulkDejaEchoue;
                    // Mode rapide :FETCh? : applicable à TOUS les types de mesure qui font des séries
                    // (Fréquence, Stabilité, FreqAvant/Final). Compatible avec l'écriture live Excel.
                    // Gain ~3× vs :READ? classique sur le 53131A. Désactivé pour Interval (1 mesure)
                    // et les types tachy/strobo (chemins de mesure différents).
                    bool typeAutoriseModeRapide = mesure.TypeMesure == TypeMesure.Stabilite
                        || mesure.TypeMesure == TypeMesure.Frequence
                        || mesure.TypeMesure == TypeMesure.FreqAvantInterv
                        || mesure.TypeMesure == TypeMesure.FreqFinale;
                    bool modeRapide = !modeBulk
                        && !string.IsNullOrEmpty(cmdFetchUsed)
                        && typeAutoriseModeRapide;

                    // Délai inter-fetch :
                    //   - 0 si la commande est bloquante (l'instrument fait l'attente)
                    //   - Sinon : gate − 100 ms (plancher 5 ms). Le write+read GPIB prend
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
                            // Amorçage adaptatif : on tente immédiatement avec un délai court
                            // (gate ms ou 50 ms minimum). Si le 1er fetch retourne 0 (erreur
                            // GPIB 230 "Data stale"), on retry en augmentant le délai.
                            // Évite de payer 200 ms × N gates quand l'instrument répond du
                            // premier coup (cas normal après une 1ère gate qui a marché).
                            // Délai = gate + 200 ms (garantit que la 1ère mesure interne est
                            // complète avant le 1er fetch — évite le retry sur data stale).
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

                    if (modeBulk)
                    {
                        // ----- Mode BULK : 1 seul aller-retour GPIB pour N mesures -----
                        // Timeout court (2 s) pour détecter rapidement les commandes non reconnues
                        // — au lieu d'attendre 5 s × N gates de pénalité.
                        _driver.DefinirTimeout(appareil.Adresse, 2000);

                        bool bulkOk = false;
                        try
                        {
                            var lot = await appareil.MesurerEnLotAsync(
                                _driver, appareil.CommandeMesureMultiple, mesure.NbMesures, ct);
                            int nbRecu = lot.Count;

                            // On considère le bulk OK si on a reçu au moins 80 % des mesures
                            // attendues. Sinon (instrument refuse la commande, retour vide,
                            // erreur SCPI), on bascule sur le mode :FETCh? rapide.
                            if (nbRecu >= mesure.NbMesures * 0.8)
                            {
                                JournalLog.Info(CategorieLog.Mesure, "BULK_FETCH",
                                    $"Lot reçu : {nbRecu}/{mesure.NbMesures} valeurs en 1 transaction GPIB.");

                                var tsBase = DateTime.Now.AddMilliseconds(-(gateSecondes * 1000) * mesure.NbMesures);
                                for (int i = 0; i < mesure.NbMesures; i++)
                                {
                                    double val = i < nbRecu ? lot[i] : 0.0;
                                    var ts = tsBase.AddMilliseconds(gateSecondes * 1000 * i);
                                    valeurs.Add(val);
                                    if (ecritureBatch) bufferBatch!.Add((ts, val));
                                    else await ExcelInteropHost.Instance.EcrireValeurLiveAsync(i, val, ts);
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

                            // Détection doublon : si même valeur que la précédente et qu'on
                            // n'est pas en mode bloquant, on attend gate/2 puis retry.
                            if (!fetchBloquant && i > 0 && val == valPrec && doublonsConsecutifs < 3)
                            {
                                doublonsConsecutifs++;
                                totalDoublons++;
                                int retryDelay = Math.Max(10, (int)(gateSecondes * 1000 * 0.5));
                                await Task.Delay(retryDelay, ct);
                                i--;        // retry l'index courant
                                continue;
                            }
                            doublonsConsecutifs = 0;
                            valPrec = val;
                        }
                        else
                        {
                            val = await appareil.MesurerAsync(_driver, mesure, ct);
                        }
                        var ts = DateTime.Now;
                        valeurs.Add(val);

                        if (ecritureBatch)
                        {
                            bufferBatch!.Add((ts, val));
                        }
                        else
                        {
                            await ExcelInteropHost.Instance.EcrireValeurLiveAsync(i, val, ts);
                        }

                        // En mode batch (Stabilité), on throttle le progress à 1 sur 5 + dernière
                        // mesure. Chaque Report() trigge un re-render du TextBox UI sur le thread
                        // dispatcher, ce qui prend 200-500 ms par update quand InformationsGenerales
                        // commence à faire plusieurs Ko (concaténation à chaque ligne) — donc 30
                        // updates = 6-15 s d'overhead pur UI sur une boucle qui devrait durer ~1 s.
                        bool reporter = !ecritureBatch || (i % 5 == 0) || (i == mesure.NbMesures - 1);
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
                    //   - Mode invisible : ClosedXML est encore ouvert en mémoire, on écrit direct
                    //   - Mode visible : on ferme Interop et on rouvre via ClosedXML
                    if (!excelInvisible)
                    {
                        await ExcelInteropHost.Instance.FermerClasseurActifAsync();
                        Perf("Interop.FermerClasseurActif");
                        await _excel.RouvrirClasseurAsync();
                        Perf("ClosedXML.RouvrirClasseurAsync");
                    }
                    await _excel.EcrireStatsAsync(valeurs);
                    Perf("ClosedXML.EcrireStatsAsync");

                    if (mesure.TypeMesure == TypeMesure.Frequence
                        || mesure.TypeMesure == TypeMesure.FreqAvantInterv
                        || mesure.TypeMesure == TypeMesure.FreqFinale)
                    {
                        await _excel.MettreAJourRecapFreqAsync(mesure);
                        Perf("ClosedXML.MettreAJourRecapFreqAsync");
                    }
                    else if (mesure.TypeMesure == TypeMesure.Stabilite)
                    {
                        await _excel.MettreAJourRecapStabAsync(mesure);
                        Perf("ClosedXML.MettreAJourRecapStabAsync");
                    }

                    progress?.Report(new ProgressionMesure
                    {
                        Message = nbIterations > 1
                            ? $"Gate {g + 1}/{nbIterations} — finalisation..."
                            : "Finalisation du rapport Excel...",
                        EtapeActuelle = etape,
                        EtapesTotales = totalEtapes
                    });
                    await _excel.SauvegarderFinalAsync();
                    Perf("ClosedXML.SauvegarderFinalAsync");

                    // Mode visible : on ferme ClosedXML et on rouvre Excel à chaque gate pour
                    // que l'utilisateur voie l'avancement.
                    // Mode invisible : on ne ferme/ouvre rien — ClosedXML reste prêt pour la
                    // gate suivante. À la dernière gate, on ferme et on ouvre Excel UNE fois
                    // pour montrer le résultat final à l'utilisateur.
                    if (!excelInvisible)
                    {
                        _excel.FermerExcel();
                        Perf("ClosedXML.FermerExcel (final)");
                        await ExcelInteropHost.Instance.OuvrirEtAfficherAsync(cheminFichier, nomFeuille);
                        Perf("Interop.OuvrirEtAfficherAsync (final)");
                    }
                    else if (isDernier)
                    {
                        _excel.FermerExcel();
                        Perf("ClosedXML.FermerExcel (final)");
                        await ExcelInteropHost.Instance.OuvrirEtAfficherAsync(cheminFichier, nomFeuille);
                        Perf("Interop.OuvrirEtAfficherAsync (1 seule fois à la fin du balayage)");
                    }

                    // À la dernière itération d'un balayage de stabilité, on ajoute le graphe
                    // de stabilité dans la Récap. (3 séries : Écart type / Valeurs Maxi / Mini
                    // en fonction du Temps de porte). Idempotent côté Interop, et on save
                    // pour persister le graphe dans le fichier final.
                    // Le graphe de stabilité est désormais embarqué directement dans le template
                    // METROLOGO_Stab.xltm (extrait de Stab1.xls historique avec toutes ses
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

                result.Succes = true;
            }
            catch (OperationCanceledException)
            {
                result.Erreur = "Mesure annulée par l'utilisateur.";
                JournalLog.Warn(CategorieLog.Mesure, "Execute", result.Erreur);
            }
            catch (Exception ex)
            {
                result.Erreur = ex.Message;
                JournalLog.Erreur(CategorieLog.Mesure, "Execute",
                    $"Échec de la mesure : {ex.Message}", new { ex.GetType().Name, ex.StackTrace });
            }
            finally
            {
                _excel.FermerExcel();
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
        private static (AppareilIEEE?, string?) ResolverDepuisCatalogue(string idModele)
        {
            var modele = CatalogueAppareilsService.Instance.Modeles
                .FirstOrDefault(m => m.Id == idModele);
            if (modele == null)
            {
                return (null, $"Le modèle catalogue « {idModele} » est introuvable. "
                    + "Le catalogue a peut-être été modifié — rouvrez la Configuration.");
            }

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
