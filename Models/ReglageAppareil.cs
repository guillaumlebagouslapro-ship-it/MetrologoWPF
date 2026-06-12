using System.Collections.Generic;

namespace Metrologo.Models
{
    /// <summary>Nature du réglage, détermine l'affichage dans la fenêtre Configuration.</summary>
    public enum TypeReglage
    {
        /// <summary>Liste d'options exclusives (ComboBox).</summary>
        Choix,

        /// <summary>Valeur numérique saisie par l'utilisateur, insérée dans un template SCPI avec {0}.</summary>
        Numerique
    }

    /// <summary>
    /// Réglage configurable d'un appareil, affiché dans la fenêtre Configuration (ComboBox pour
    /// Choix, saisie libre pour Numerique). Tout vit dans le JSON catalogue, donc on ajoute des
    /// réglages sans toucher au code.
    /// </summary>
    public class ReglageAppareil
    {
        /// <summary>Libellé du réglage (ex: "Impédance Voie A", "Trigger Voie A").</summary>
        public string Nom { get; set; } = string.Empty;

        /// <summary>Type de réglage. Défaut Choix pour la rétrocompat JSON.</summary>
        public TypeReglage Type { get; set; } = TypeReglage.Choix;

        /// <summary>
        /// Pour Choix : options proposées (la première sélectionnée par défaut).
        /// Pour Numerique : une seule option dont CommandeScpi est le template avec {0}.
        /// </summary>
        public List<OptionReglage> Options { get; set; } = new();

        /// <summary>Unité affichée à côté du champ numérique (ex: "V"). Ignorée pour Choix.</summary>
        public string Unite { get; set; } = string.Empty;

        /// <summary>Valeur numérique par défaut (ex: "0.5"). Ignorée pour Choix.</summary>
        public string ValeurDefaut { get; set; } = string.Empty;

        /// <summary>
        /// true = réglage auto-sélectionné selon le contexte (TypeMesure + VoieActive), masqué dans
        /// la fenêtre Configuration user. Le code parcourt les Options et prend celle dont le libellé
        /// matche : "FREQ Voie A/B/C" selon la voie en fréquence, "TIAB" pour Interval, "CONT" pour
        /// Stabilite, "AUTO" en fallback. Typique du 53230A ("Mode de mesure", "Résolution").
        /// </summary>
        public bool Auto { get; set; } = false;
    }

    /// <summary>Option d'un réglage : libellé affiché + commande SCPI correspondante.</summary>
    public class OptionReglage
    {
        /// <summary>Libellé affiché dans la ComboBox (ex: "50 Ω").</summary>
        public string Libelle { get; set; } = string.Empty;

        /// <summary>Commande SCPI envoyée quand cette option est sélectionnée (ex: "INP:IMP 50").</summary>
        public string CommandeScpi { get; set; } = string.Empty;

        /// <summary>
        /// Pour un réglage Auto : nom du TypeMesure requis pour prendre cette option (ex: "Stabilite",
        /// "Interval"). Vide = tous les types. Si plusieurs options matchent, la plus spécifique gagne.
        /// </summary>
        public string QuandType { get; set; } = string.Empty;

        /// <summary>Pour un réglage Auto : voie active requise ("A", "B", "C"). Vide = toutes les voies.</summary>
        public string QuandVoie { get; set; } = string.Empty;

        public override string ToString() => Libelle;
    }
}
