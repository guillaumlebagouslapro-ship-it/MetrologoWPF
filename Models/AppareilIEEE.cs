using System.Collections.Generic;

namespace Metrologo.Models
{
    // config d'un temps de porte (gate)
    public class GateConfig
    {
        public string Libelle { get; set; } = string.Empty;
        public string Commande { get; set; } = string.Empty;
        public double ValeurSecondes { get; set; }
    }

    public class AppareilIEEE
    {
        public string Nom { get; set; } = string.Empty;
        public int Adresse { get; set; }

        // terminaisons et entête de réponse
        public int WriteTerm { get; set; }
        public int ReadTerm { get; set; }
        public int TailleHeaderReponse { get; set; }

        // commandes SCPI ou spécifiques à l'appareil
        public string ChaineInit { get; set; } = string.Empty;
        public string ConfEntree { get; set; } = string.Empty;
        public string ExeMesure { get; set; } = string.Empty;
        public string Monocoup { get; set; } = string.Empty;

        /// <summary>Commande de mesure en lot (placeholder {N}) : N mesures en un seul aller-retour GPIB.
        /// Vide = fallback boucle :FETCh? puis :READ?.</summary>
        public string CommandeMesureMultiple { get; set; } = string.Empty;

        /// <summary>Commande SCPI bloquante jusqu'à la prochaine mesure dispo.
        /// Si non vide, évite le Task.Delay entre fetches en boucle rapide.</summary>
        public string CommandeFetchFresh { get; set; } = string.Empty;

        // service request (SRQ)
        public bool GereSRQ { get; set; }
        public string SRQOn { get; set; } = string.Empty;
        public string SRQOff { get; set; } = string.Empty;

        /// <summary>
        /// Vérification d'arming après gate (commandes :FREQ:ARM:* spécifiques 53131A). true uniquement
        /// pour les modèles qui supportent cette syntaxe ; défaut false pour éviter le -113 Undefined header
        /// sur les compteurs modernes (53230A, SR620...).
        /// </summary>
        public bool VerifArmingActive { get; set; } = false;

        /// <summary>
        /// Mode rapide :INIT:CONT ON + boucle :FETCh? dans l'orchestrator. À false pour les compteurs
        /// modernes (53230A) qui n'aiment pas cette séquence : on retombe alors sur :READ? par mesure
        /// ou sur le bulk si configuré. Défaut true (rétrocompat 53131A).
        /// </summary>
        public bool ModeRapideActif { get; set; } = true;

        /// <summary>
        /// Commande qui lance une acquisition gap-free sans renvoyer les valeurs (placeholder {N}).
        /// Avec CommandeFetchFresh, active le streaming : acquisition gap-free puis lecture valeur
        /// par valeur. Vide = bulk classique.
        /// </summary>
        public string CommandeBulkInit { get; set; } = string.Empty;

        // remplace les 3 TStringList de Delphi : clé = index de gate, valeur = GateConfig
        public Dictionary<int, GateConfig> Gates { get; set; } = new Dictionary<int, GateConfig>();
    }
}