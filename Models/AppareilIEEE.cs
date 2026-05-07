using System.Collections.Generic;

namespace Metrologo.Models
{
    // Petite classe pour regrouper proprement les configurations de "Gate" (Temps de porte)
    public class GateConfig
    {
        public string Libelle { get; set; } = string.Empty;
        public string Commande { get; set; } = string.Empty;
        public double ValeurSecondes { get; set; }
    }

    // Modèle principal de l'appareil
    public class AppareilIEEE
    {
        public string Nom { get; set; } = string.Empty;
        public int Adresse { get; set; }

        // Configuration de la communication
        public int WriteTerm { get; set; }
        public int ReadTerm { get; set; }
        public int TailleHeaderReponse { get; set; }

        // Commandes (SCPI ou spécifiques à l'appareil)
        public string ChaineInit { get; set; } = string.Empty;
        public string ConfEntree { get; set; } = string.Empty;
        public string ExeMesure { get; set; } = string.Empty;
        public string Monocoup { get; set; } = string.Empty;

        /// <summary>
        /// Commande de mesure en lot (avec placeholder <c>{N}</c>). Si non vide, l'orchestrator
        /// l'utilise pour récupérer N mesures en un seul aller-retour GPIB. Vide = fallback
        /// sur :FETCh? boucle puis :READ?.
        /// </summary>
        public string CommandeMesureMultiple { get; set; } = string.Empty;

        /// <summary>
        /// Commande SCPI qui bloque jusqu'à ce qu'une nouvelle mesure soit disponible.
        /// Si non vide, supprime le Task.Delay entre fetches dans la boucle rapide.
        /// </summary>
        public string CommandeFetchFresh { get; set; } = string.Empty;

        // Gestion des interruptions (Service Request)
        public bool GereSRQ { get; set; }
        public string SRQOn { get; set; } = string.Empty;
        public string SRQOff { get; set; } = string.Empty;

        /// <summary>
        /// Active la vérification d'arming après gate (commandes <c>:FREQ:ARM:*</c> spécifiques
        /// 53131A). À <c>true</c> uniquement pour les modèles qui supportent cette syntaxe ;
        /// défaut <c>false</c> pour rester universel et ne pas générer de <c>-113 Undefined
        /// header</c> sur les compteurs modernes (53230A, SR620, etc.).
        /// </summary>
        public bool VerifArmingActive { get; set; } = false;

        /// <summary>
        /// Active le mode rapide <c>:INIT:CONT ON</c> + boucle <c>:FETCh?</c> dans l'orchestrator.
        /// À <c>false</c> pour les compteurs modernes (53230A) qui n'aiment pas cette séquence —
        /// l'orchestrator tombe alors sur le mode classique <c>:READ?</c>/mesure ou sur le bulk
        /// (<see cref="CommandeMesureMultiple"/>) si configuré. Défaut <c>true</c> (rétrocompat 53131A).
        /// </summary>
        public bool ModeRapideActif { get; set; } = true;

        /// <summary>
        /// Commande pour lancer une acquisition gap-free sans renvoyer les valeurs (placeholder
        /// <c>{N}</c>). Combinée à <see cref="CommandeFetchFresh"/>, active le mode streaming :
        /// l'orchestrator lance l'acquisition en gap-free puis lit valeur par valeur. Vide =
        /// pas de streaming (l'orchestrator retombe sur le bulk classique).
        /// </summary>
        public string CommandeBulkInit { get; set; } = string.Empty;

        // Remplacement des 3 TStringList de Delphi par un Dictionnaire C# propre
        // La clé (int) sera l'index de la Gate, et la valeur sera l'objet GateConfig
        public Dictionary<int, GateConfig> Gates { get; set; } = new Dictionary<int, GateConfig>();
    }
}