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

        /// <summary>Réglages dynamiques de la fenêtre Configuration (impédance, filtre,
        /// atténuation...). Chaque réglage porte ses options et commandes SCPI — UI
        /// extensible sans recompilation.</summary>
        public List<ReglageAppareil> Reglages { get; set; } = new();

        // stocké en UTC, ToLocalTime() pour l'affichage
        public DateTime DateCreation { get; set; } = DateTime.UtcNow;
        public string CreePar { get; set; } = string.Empty;

        /// <summary>
        /// Nombre de voies physiques (1, 2 ou 3, la 3e étant souvent une voie HF). Pilote l'affichage
        /// des sections de réglages. Défaut 2 pour les modèles historiques (53131A, 53230A) qui ont A + B.
        /// </summary>
        public int NbVoies { get; set; } = 2;

        /// <summary>Vérifie si cet IDN correspond à ce modèle. Le MODÈLE doit correspondre
        /// EXACTEMENT (insensible casse/espaces) ; le fabricant reste tolérant (les chaînes IDN
        /// varient : « Agilent Technologies », « Keysight Technologies », « HEWLETT-PACKARD »…).</summary>
        public bool Correspond(string? fabricant, string? modele)
        {
            string fab = (fabricant ?? string.Empty).Trim().ToUpperInvariant();
            string mod = (modele ?? string.Empty).Trim().ToUpperInvariant();
            string fabP = (FabricantIdn ?? string.Empty).Trim().ToUpperInvariant();
            string modP = (ModeleIdn ?? string.Empty).Trim().ToUpperInvariant();

            // Fabricant : vide = critère ignoré ; sinon « contient » (tolérant aux variantes).
            bool fabOk = string.IsNullOrEmpty(fabP) || fab.Contains(fabP);

            // Modèle : correspondance EXACTE, et JAMAIS sur un motif vide.
            // Avant, un « contains » faisait correspondre un 53220A à une fiche 53230A, et une
            // fiche au modèle vide matchait TOUT appareil — d'où des appareils « reconnus » alors
            // qu'ils n'avaient jamais été enregistrés, et l'impossibilité de les enregistrer à part.
            // Avec l'égalité exacte, chaque modèle est indépendant (53220A ≠ 53230A) : un appareil
            // n'est reconnu que si SON modèle précis figure au catalogue.
            bool modOk = !string.IsNullOrEmpty(modP) && mod == modP;
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
        /// Commande bulk N mesures côté hardware, renvoyées en CSV. {N} = nombre de mesures
        /// (ex 53131A : ":SAMP:COUN {N};:READ:ARR? {N}"). Si renseignée, l'orchestrator passe
        /// en bulk au lieu de boucler sur :FETCh? (10 ms x 30 : ~0,5 s contre ~6 s).
        /// Vide = fallback :INIT:CONT ON + :FETCh?, puis :READ?. Stratégie entièrement pilotée
        /// par le catalogue, rien de codé en dur.
        /// </summary>
        public string CommandeMesureMultiple { get; set; } = string.Empty;

        /// <summary>Commande bloquante qui renvoie la prochaine mesure (évite le Task.Delay
        /// entre fetches). Ex 53131A : ":DATA:FRESh:FREQ?" selon firmware.
        /// Vide = fallback :FETCh? + Task.Delay (gate × 0,5) + détection doublons.</summary>
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

        /// <summary>Vérif de l'arming après envoi du gate (:FREQ:ARM:STOP:SOUR? + :TIM? + :SYST:ERR?).
        /// 53130A/compatibles uniquement : les autres (53230A, SR620...) renvoient -113 Undefined header
        /// sur ces commandes. Défaut false.</summary>
        public bool VerifArmingActive { get; set; } = false;

        /// <summary>
        /// Active :INIT:CONT ON + boucle :FETCh? (ok pour 53131A). À désactiver sur 53230A et
        /// compteurs modernes : leur :INIT:CONT ON est un "front panel running", la combinaison
        /// avec :FETCh? donne -113 ou +230 Data corrupt. Ceux-ci utilisent SAMP:COUN N + READ?
        /// via CommandeMesureMultiple. Défaut true (rétrocompat 53131A).
        /// </summary>
        public bool ModeRapideActif { get; set; } = true;

        /// <summary>Lance une acquisition gap-free de N mesures sans renvoyer les valeurs ({N} =
        /// placeholder). Mode streaming pour gates longues (>= 1 s) : bulk lu une mesure à la fois
        /// via CommandeFetchFresh. Ex 53230A : :SENS:FREQ:MODE CONT;:SAMP:COUN {N};:INIT:IMM.
        /// Vide ou CommandeFetchFresh vide = bulk classique en une transaction GPIB.</summary>
        public string CommandeBulkInit { get; set; } = string.Empty;

        // -------- appareils legacy (pas de *IDN?) --------

        /// <summary>Appareil sans *IDN? (EIP 545, Racal-Dana 1996, Stanford SR620).
        /// Résolution par AdresseFixeParDefaut au lieu du match IDN. Défaut false.</summary>
        public bool Legacy { get; set; } = false;

        /// <summary>
        /// Adresse GPIB par défaut d'un appareil legacy / en mode adresses fixes. Pré-remplie dans
        /// l'UI (éditable selon le banc), ignorée pour les appareils détectés par IDN.
        /// </summary>
        public int AdresseFixeParDefaut { get; set; } = 0;

        /// <summary>Map slot UI (0=10 ms … 12=100 s) → commande de gate discrète, quand
        /// CommandeGate ne suffit pas. Ex Stanford slot 6 = "armm5;size1E0", EIP = "R0/R1/R2".
        /// Prioritaire sur CommandeGate dans CatalogueAdapter. Vide = template inchangé.</summary>
        public Dictionary<int, string> CommandesGateParSlot { get; set; } = new();

        /// <summary>Templates SCPI pour la mesure d'intervalle. Si <see cref="CommandesIntervalle.Actif"/>
        /// est false (défaut), le panneau est masqué et rien n'est envoyé — évite d'envoyer des
        /// commandes 53230A à un Stanford.</summary>
        public CommandesIntervalle Intervalle { get; set; } = new();

        /// <summary>Copie profonde : pour cloner les réglages d'un modèle vers un autre (ex.
        /// 53230A → 53220A) SANS oublier les champs « cachés » du formulaire (ModeRapideActif,
        /// VerifArmingActive, CommandeBulkInit…) ni partager les références mutables.</summary>
        public ParametresIeee Cloner()
        {
            var c = (ParametresIeee)MemberwiseClone();
            c.CommandesGateParSlot = new Dictionary<int, string>(CommandesGateParSlot);
            return c;
        }
    }

    /// <summary>
    /// Templates SCPI pour mesure d'intervalle, stockés par appareil dans le catalogue.
    /// Placeholders : {V}=voie (1/2), {C}=couplage (AC|DC), {Z}=impédance (50|1E6),
    /// {S}=seuil (V), {P}=pente (POS|NEG), {T}=hold-off (s).
    /// Template vide = commande ignorée. Chaînage possible avec ";:".
    /// </summary>
    public class CommandesIntervalle
    {
        /// <summary>L'appareil gère l'intervalle piloté logiciel. false = panneau masqué, rien envoyé.</summary>
        public bool Actif { get; set; }

        public string Conf1Voie { get; set; } = string.Empty;       // ex: "CONF:TINT (@1)"
        public string Conf2Voies { get; set; } = string.Empty;      // ex: "CONF:TINT (@1),(@2)"
        public string Couplage { get; set; } = string.Empty;        // ex: "INP{V}:COUP {C}"
        public string Impedance { get; set; } = string.Empty;       // ex: "INP{V}:IMP {Z}"
        public string SeuilStart { get; set; } = string.Empty;      // ex: "INP{V}:LEV1 {S}"
        public string PenteStart { get; set; } = string.Empty;      // ex: "INP{V}:SLOP1 {P}"
        public string SeuilStop1Voie { get; set; } = string.Empty;  // ex: "INP1:LEV2 {S}" (mode 1 voie)
        public string PenteStop1Voie { get; set; } = string.Empty;  // ex: "INP1:SLOP2 {P}" (mode 1 voie)
        public string Holdoff { get; set; } = string.Empty;         // ex chaîné: "TINT:GATE:SOUR ADV;:SENS:GATE:STOP:HOLD:SOUR TIME;:SENS:GATE:STOP:HOLD:TIME {T};:SENS:GATE:STOP:SOUR IMM"

        /// <summary>Jeu de commandes standard du Keysight 53230A, pour pré-remplir le formulaire.</summary>
        public static CommandesIntervalle Defaut53230A() => new()
        {
            Actif = true,
            Conf1Voie = "CONF:TINT (@1)",
            Conf2Voies = "CONF:TINT (@1),(@2)",
            Couplage = "INP{V}:COUP {C}",
            Impedance = "INP{V}:IMP {Z}",
            SeuilStart = "INP{V}:LEV1 {S}",
            PenteStart = "INP{V}:SLOP1 {P}",
            SeuilStop1Voie = "INP1:LEV2 {S}",
            PenteStop1Voie = "INP1:SLOP2 {P}",
            Holdoff = "TINT:GATE:SOUR ADV;:SENS:GATE:STAR:SOUR IMM;:SENS:GATE:STOP:HOLD:SOUR TIME;:SENS:GATE:STOP:HOLD:TIME {T};:SENS:GATE:STOP:SOUR IMM"
        };
    }
}
