using System;
using System.Collections.Generic;

namespace Metrologo.Models
{
    /// <summary>
    /// Modèle d'appareil enregistré dans le catalogue local.
    /// Identifié par un ID stable (slug) et une signature IDN servant au matching lors du scan.
    /// </summary>
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
        /// Réglages dynamiques affichés dans la fenêtre Configuration et envoyés à l'appareil
        /// lors de la validation (ex: impédance, filtre, atténuation). Permet d'étendre la UI
        /// sans recompiler : chaque réglage porte ses options + leurs commandes SCPI.
        /// </summary>
        public List<ReglageAppareil> Reglages { get; set; } = new();

        public DateTime DateCreation { get; set; } = DateTime.Now;
        public string CreePar { get; set; } = string.Empty;

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
        /// Modèle de commande pour programmer le gate time. <c>{0}</c> sera remplacé
        /// par le temps en secondes. Ex: <c>":FREQ:ARM:STOP:TIM {0}"</c>.
        /// Vide = pas de gate programmable (ou l'appareil n'a pas cette notion).
        /// </summary>
        public string CommandeGate { get; set; } = string.Empty;

        /// <summary>
        /// Commande de mesure en lot — si l'appareil sait faire N mesures côté hardware
        /// et les retourner d'un coup en CSV. <c>{N}</c> sera remplacé par le nombre de
        /// mesures. Ex 53131A : <c>":SAMP:COUN {N};:READ:ARR? {N}"</c>.
        ///
        /// Si non vide, l'orchestrator utilise ce mode bulk au lieu du :FETCh? boucle —
        /// gain massif sur les boucles courtes (10 ms × 30 → ~0,5 s vs ~6 s en :FETCh?
        /// boucle), car les mesures se font côté instrument sans aller-retour GPIB.
        ///
        /// Vide = fallback automatique sur :INIT:CONT ON + :FETCh? (mode rapide), puis
        /// sur :READ? si Fetch n'est pas dérivable. Pas de comportement spécifique à un
        /// modèle en dur côté C# — la stratégie est entièrement décidée par le catalogue.
        /// </summary>
        public string CommandeMesureMultiple { get; set; } = string.Empty;

        /// <summary>
        /// Commande SCPI qui retourne la **prochaine** mesure (et bloque jusqu'à ce
        /// qu'elle soit dispo) au lieu de la dernière déjà lue. Permet d'éviter le
        /// Task.Delay entre deux fetches dans la boucle de mesures rapides.
        ///
        /// Ex 53131A : <c>":DATA:FRESh:FREQ?"</c> ou similaire selon firmware. Vide =
        /// fallback sur :FETCh? + Task.Delay = gate × 0,5 + détection doublons.
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
    }
}
