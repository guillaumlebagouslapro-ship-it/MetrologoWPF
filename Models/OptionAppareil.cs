namespace Metrologo.Models
{
    /// <summary>
    /// Entrée unifiée pour la dropdown "Fréquencemètre" de la fenêtre Configuration.
    /// Peut représenter soit un type connu du catalogue (Stanford/Racal/EIP), soit un appareil
    /// détecté sur le bus qui ne fait pas partie du catalogue.
    /// </summary>
    public class OptionAppareil
    {
        /// <summary>Libellé affiché dans la dropdown.</summary>
        public string Libelle { get; init; } = string.Empty;

        /// <summary>Type catalogue (Stanford/Racal/EIP). Null si appareil détecté inconnu.</summary>
        public TypeAppareilIEEE? Type { get; init; }

        /// <summary>Détails de détection sur le bus. Null si jamais détecté.</summary>
        public AppareilDetecte? Detecte { get; init; }

        public bool EstDetecte => Detecte != null;
        public bool EstConnu => Type.HasValue;
        public string AdresseAffiche => Detecte?.AdresseCourte ?? "non connecté";

        public override string ToString() => Libelle;
    }
}
