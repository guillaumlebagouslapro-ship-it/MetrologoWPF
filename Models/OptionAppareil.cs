namespace Metrologo.Models
{
    /// <summary>
    /// Entrée de la dropdown "Fréquencemètre" de la fenêtre Configuration.
    /// Représente un appareil détecté sur le bus GPIB, éventuellement associé à un modèle du catalogue.
    /// </summary>
    public class OptionAppareil
    {
        /// <summary>Libellé affiché dans la dropdown.</summary>
        public string Libelle { get; init; } = string.Empty;

        /// <summary>Détails de détection sur le bus. Null si jamais détecté.</summary>
        public AppareilDetecte? Detecte { get; init; }

        public bool EstDetecte => Detecte != null;
        public bool EstConnu => Detecte?.ModeleReconnu != null;
        public string AdresseAffiche => Detecte?.AdresseCourte ?? "non connecté";

        public override string ToString() => Libelle;
    }
}
