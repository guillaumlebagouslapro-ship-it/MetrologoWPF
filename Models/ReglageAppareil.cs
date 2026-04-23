using System.Collections.Generic;

namespace Metrologo.Models
{
    /// <summary>Nature du réglage — détermine comment la UI Configuration l'affiche.</summary>
    public enum TypeReglage
    {
        /// <summary>Liste d'options mutuellement exclusives (ComboBox).</summary>
        Choix,

        /// <summary>Valeur numérique saisie par l'utilisateur, insérée dans un template SCPI contenant <c>{0}</c>.</summary>
        Numerique
    }

    /// <summary>
    /// Réglage configurable d'un appareil exposé dans la fenêtre Configuration.
    /// Type <see cref="TypeReglage.Choix"/> : liste d'options → ComboBox.
    /// Type <see cref="TypeReglage.Numerique"/> : saisie libre formatée dans un template SCPI.
    ///
    /// Permet d'ajouter des réglages sans toucher au code — tout vit dans le JSON catalogue.
    /// </summary>
    public class ReglageAppareil
    {
        /// <summary>Libellé du réglage (ex: "Impédance Voie A", "Trigger Voie A").</summary>
        public string Nom { get; set; } = string.Empty;

        /// <summary>Type de réglage — par défaut <c>Choix</c> pour rétrocompatibilité JSON.</summary>
        public TypeReglage Type { get; set; } = TypeReglage.Choix;

        /// <summary>
        /// Pour Type=<c>Choix</c> : options proposées (la première est sélectionnée par défaut).
        /// Pour Type=<c>Numerique</c> : une seule option dont <c>CommandeScpi</c> est le template avec <c>{0}</c>.
        /// </summary>
        public List<OptionReglage> Options { get; set; } = new();

        /// <summary>Unité affichée à côté du champ numérique (ex: "V"). Ignorée pour Choix.</summary>
        public string Unite { get; set; } = string.Empty;

        /// <summary>Valeur numérique par défaut (ex: "0.5"). Ignorée pour Choix.</summary>
        public string ValeurDefaut { get; set; } = string.Empty;
    }

    /// <summary>
    /// Option d'un <see cref="ReglageAppareil"/> : libellé affiché + commande SCPI correspondante.
    /// </summary>
    public class OptionReglage
    {
        /// <summary>Libellé affiché dans la ComboBox (ex: "50 Ω").</summary>
        public string Libelle { get; set; } = string.Empty;

        /// <summary>Commande SCPI envoyée quand cette option est sélectionnée (ex: "INP:IMP 50").</summary>
        public string CommandeScpi { get; set; } = string.Empty;

        public override string ToString() => Libelle;
    }
}
