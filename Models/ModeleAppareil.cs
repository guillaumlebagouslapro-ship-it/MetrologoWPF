using System;
using System.Collections.Generic;

namespace Metrologo.Models
{
    /// <summary>Modèle d'appareil du catalogue local. Id stable (slug) + signature IDN pour le matching au scan.</summary>
    public class ModeleAppareil
    {
        /// <summary>Identifiant stable, généré à la création (ex: "agilent-53131a-7f3a").</summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>Nom lisible affiché dans les listes (ex: "Agilent 53131A").</summary>
        public string Nom { get; set; } = string.Empty;

        /// <summary>Motif de reconnaissance sur la chaîne IDN (insensible à la casse, matching "contains").</summary>
        public string FabricantIdn { get; set; } = string.Empty;
        public string ModeleIdn { get; set; } = string.Empty;

        /// <summary>Paramètres SCPI/IEEE pour dialoguer avec l'appareil.</summary>
        public ParametresIeee Parametres { get; set; } = new();

        /// <summary>Libellés des temps de porte disponibles (ex: "10 ms", "1 s"). Vide = pas de gate.</summary>
        public List<string> Gates { get; set; } = new();

        /// <summary>Libellés des entrées/gammes proposées dans la UI (ex: "Ch1 50Ω"). Vide = pas de choix.</summary>
        public List<string> Entrees { get; set; } = new();

        /// <summary>Couplages proposés (ex: "AC", "DC"). Vide = pas de choix.</summary>
        public List<string> Couplages { get; set; } = new();

        /// <summary>
        /// Réglages dynamiques de la fenêtre Configuration (impédance, filtre, atténuation...).
        /// Chaque réglage porte ses options et leurs commandes SCPI, donc on peut étendre la UI sans recompiler.
        /// </summary>
        public List<ReglageAppareil> Reglages { get; set; } = new();

        // stocké en UTC, ToLocalTime() pour l'affichage
        public DateTime DateCreation { get; set; } = DateTime.UtcNow;
        public string CreePar { get; set; } = string.Empty;

        /// <summary>
        /// Nombre de voies physiques (1, 2 ou 3, la 3e étant souvent une voie HF). Pilote l'affichage
        /// des sections de réglages. Défaut 2 pour les modèles historiques (53131A, 53230A) qui ont A + B.
        /// </summary>
        public int NbVoies { get; set; } = 2;

        /// <summary>Vérifie si cet IDN correspond à ce modèle (contains insensible à la casse).</summary>
        public bool Correspond(string? fabricant, string? modele)
        {
            string fab = (fabricant ?? string.Empty).ToUpperInvariant();
            string mod = (modele ?? string.Empty).ToUpperInvariant();
            string fabP = (FabricantIdn ?? string.Empty).ToUpperInvariant();
            string modP = (ModeleIdn ?? string.Empty).ToUpperInvariant();

            bool fabOk = string.IsNullOrEmpty(fabP) || fab.Contains(fabP);
            bool modOk = string.IsNullOrEmpty(modP) || mod.Contains(modP);
            return fabOk && modOk;
        }
    }

    /// <summary>Commandes et paramètres IEEE pour piloter un appareil SCPI ou historique.</summary>
    public class ParametresIeee
    {
        /// <summary>Chaîne d'initialisation envoyée au début (ex: "*RST;*CLS").</summary>
        public string ChaineInit { get; set; } = "*RST;*CLS";

        /// <summary>Configuration d'entrée envoyée après l'init (optionnelle).</summary>
        public string ConfEntree { get; set; } = string.Empty;

        /// <summary>Commande pour déclencher et lire une mesure (ex: ":READ?").</summary>
        public string ExeMesure { get; set; } = ":READ?";

        /// <summary>
        /// Template de commande pour le gate time, {0} = temps en secondes (ex: ":FREQ:ARM:STOP:TIM {0}").
        /// Vide = pas de gate programmable.
        /// </summary>
        public string CommandeGate { get; set; } = string.Empty;

        /// <summary>
        /// Commande de mesure en lot quand l'appareil sait faire N mesures côté hardware et les
        /// renvoyer d'un coup en CSV. {N} = nombre de mesures (ex 53131A : ":SAMP:COUN {N};:READ:ARR? {N}").
        /// Si non vide, l'orchestrator passe en bulk au lieu de boucler sur :FETCh?, gros gain sur les
        /// gates courtes (10 ms x 30 : ~0,5 s contre ~6 s) car plus d'aller-retour GPIB par mesure.
        /// Vide = fallback :INIT:CONT ON + :FETCh?, puis :READ?. La stratégie vient entièrement du
        /// catalogue, rien de codé en dur par modèle.
        /// </summary>
        public string CommandeMesureMultiple { get; set; } = string.Empty;

        /// <summary>
        /// Commande SCPI qui renvoie la prochaine mesure (bloquante) au lieu de la dernière déjà lue,
        /// ce qui évite le Task.Delay entre deux fetches en boucle rapide. Ex 53131A : ":DATA:FRESh:FREQ?"
        /// selon firmware. Vide = fallback :FETCh? + Task.Delay (gate x 0,5) + détection doublons.
        /// </summary>
        public string CommandeFetchFresh { get; set; } = string.Empty;

        /// <summary>WriteTerm : 0=none, 1=NL (LF), 2=EOI. Cf. Metrologo.ini.</summary>
        public int TermWrite { get; set; } = 1;

        /// <summary>ReadTerm : 10=LF (ASCII), 13=CR, 256=EOI only.</summary>
        public int TermRead { get; set; } = 10;

        /// <summary>Nombre de caractères d'entête à sauter avant le nombre (1 = pas de saut).</summary>
        public int TailleHeader { get; set; } = 1;

        public bool GereSrq { get; set; } = false;
        public string SrqOn { get; set; } = string.Empty;
        public string SrqOff { get; set; } = string.Empty;

        /// <summary>
        /// Vérif optionnelle de l'arming après envoi du gate (:FREQ:ARM:STOP:SOUR? + :TIM? + :SYST:ERR?).
        /// Spécifique au HP/Agilent 53131A et compatibles : les autres (53230A, SR620...) répondent
        /// -113 Undefined header sur ces commandes ARM, qui s'affiche à l'écran de l'instrument.
        /// Défaut false, à activer seulement si le modèle supporte la syntaxe :FREQ:ARM:* historique.
        /// </summary>
        public bool VerifArmingActive { get; set; } = false;

        /// <summary>
        /// Active le mode rapide :INIT:CONT ON + boucle :FETCh? dans l'orchestrator (ok pour le 53131A
        /// et compatibles). A désactiver pour le 53230A et les compteurs modernes : leur :INIT:CONT ON
        /// fait du "front panel running" et la combinaison avec :FETCh? donne -113 ou +230 Data corrupt,
        /// ils passent par SAMP:COUN N + READ? via CommandeMesureMultiple. Défaut true (rétrocompat
        /// 53131A) ; à false l'orchestrator retombe sur :READ? par mesure ou sur le bulk si configuré.
        /// </summary>
        public bool ModeRapideActif { get; set; } = true;

        /// <summary>
        /// Commande qui lance une acquisition gap-free de N mesures sans renvoyer les valeurs
        /// (placeholder {N}). Sert au mode streaming (bulk lu une mesure à la fois via
        /// CommandeFetchFresh) pour les gates longues (1 s et plus) où on veut du live à l'écran.
        /// Ex 53230A : :SENS:FREQ:MODE CONT;:SAMP:COUN {N};:INIT:IMM. Si vide (ou si
        /// CommandeFetchFresh est vide), pas de streaming : bulk classique en une transaction GPIB.
        /// </summary>
        public string CommandeBulkInit { get; set; } = string.Empty;

        // -------- appareils legacy (pas de *IDN?) --------

        /// <summary>
        /// true = appareil ancien qui ne répond pas à *IDN? (EIP 545, Racal-Dana 1996, Stanford SR620).
        /// Indétectable au scan GPIB : l'orchestrator le résout par AdresseFixeParDefaut au lieu d'un
        /// match IDN. Défaut false (comportement IDN inchangé pour les compteurs modernes).
        /// </summary>
        public bool Legacy { get; set; } = false;

        /// <summary>
        /// Adresse GPIB par défaut d'un appareil legacy / en mode adresses fixes. Pré-remplie dans
        /// l'UI (éditable selon le banc), ignorée pour les appareils détectés par IDN.
        /// </summary>
        public int AdresseFixeParDefaut { get; set; } = 0;

        /// <summary>
        /// Map slot UI (0=10 ms ... 12=100 s) vers une commande de gate discrète quand le template
        /// CommandeGate ne suffit pas. Ex Stanford slot 6 = "armm5;size1E0", Racal slot 6 = "GA1E0",
        /// EIP slots 0/3/6 = "R2"/"R1"/"R0". Si non vide, prioritaire sur CommandeGate dans
        /// CatalogueAdapter. Vide = comportement template inchangé.
        /// </summary>
        public Dictionary<int, string> CommandesGateParSlot { get; set; } = new();
    }
}
