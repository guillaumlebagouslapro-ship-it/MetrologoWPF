using System.Collections.Generic;

namespace Metrologo.Models
{
    /// <summary>Nature du réglage : c'est ce qui décide de la façon de l'afficher dans la fenêtre Configuration.</summary>
    public enum TypeReglage
    {
        /// <summary>Une liste d'options exclusives, présentée en ComboBox.</summary>
        Choix,

        /// <summary>Une valeur numérique saisie par l'utilisateur, qu'on injecte dans un template SCPI à la place du {0}.</summary>
        Numerique
    }

    /// <summary>
    /// Un réglage configurable d'un appareil, tel qu'il apparaît dans la fenêtre Configuration
    /// (ComboBox pour un Choix, saisie libre pour un Numerique). Tout est décrit dans le JSON du
    /// catalogue, ce qui permet d'ajouter des réglages sans retoucher au code.
    /// </summary>
    public class ReglageAppareil
    {
        /// <summary>Libellé du réglage, par ex. "Impédance Voie A" ou "Trigger Voie A".</summary>
        public string Nom { get; set; } = string.Empty;

        /// <summary>Type de réglage. On part sur Choix par défaut, pour rester compatible avec les anciens JSON.</summary>
        public TypeReglage Type { get; set; } = TypeReglage.Choix;

        /// <summary>
        /// Pour un Choix : la liste des options proposées (la première est sélectionnée d'office).
        /// Pour un Numerique : une seule option, dont la CommandeScpi sert de template avec son {0}.
        /// </summary>
        public List<OptionReglage> Options { get; set; } = new();

        /// <summary>Unité affichée juste à côté du champ numérique (par ex. "V"). Sans effet pour un Choix.</summary>
        public string Unite { get; set; } = string.Empty;

        /// <summary>Valeur numérique proposée par défaut (par ex. "0.5"). Sans effet pour un Choix.</summary>
        public string ValeurDefaut { get; set; } = string.Empty;

        /// <summary>
        /// true = le réglage se choisit tout seul selon le contexte (TypeMesure + VoieActive) et reste
        /// caché dans la fenêtre Configuration côté user. Le code parcourt les Options et retient celle
        /// dont le libellé colle : "FREQ Voie A/B/C" selon la voie en fréquence, "TIAB" pour Interval,
        /// "CONT" pour Stabilite, et "AUTO" en dernier recours. Typiquement le cas du 53230A ("Mode de
        /// mesure", "Résolution").
        /// </summary>
        public bool Auto { get; set; } = false;
    }

    /// <summary>Une option d'un réglage : le libellé affiché et la commande SCPI qui va avec.</summary>
    public class OptionReglage
    {
        /// <summary>Libellé tel qu'il apparaît dans la ComboBox (par ex. "50 Ω").</summary>
        public string Libelle { get; set; } = string.Empty;

        /// <summary>Commande SCPI envoyée quand on choisit cette option (par ex. "INP:IMP 50").</summary>
        public string CommandeScpi { get; set; } = string.Empty;

        /// <summary>
        /// Pour un réglage Auto : le TypeMesure qu'il faut pour que cette option soit retenue (par ex.
        /// "Stabilite", "Interval"). Vide = valable pour tous les types. Si plusieurs options collent,
        /// c'est la plus spécifique qui l'emporte.
        /// </summary>
        public string QuandType { get; set; } = string.Empty;

        /// <summary>Pour un réglage Auto : la voie active attendue ("A", "B", "C"). Vide = toutes les voies.</summary>
        public string QuandVoie { get; set; } = string.Empty;

        public override string ToString() => Libelle;
    }
}
