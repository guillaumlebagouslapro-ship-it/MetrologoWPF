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

        // ---------------- mode adresses fixes (appareils legacy non détectés) ----------------

        /// <summary>
        /// Modèle catalogue legacy associé en mode adresses fixes (l'appareil n'est pas sur le bus,
        /// on le pilote à une adresse saisie). Null en mode scan.
        /// </summary>
        public ModeleAppareil? ModeleFixe { get; init; }

        public bool EstFixe => ModeleFixe != null;

        /// <summary>
        /// Adresse GPIB éditable (mode adresses fixes), pré-remplie depuis le profil. Settable :
        /// liée en TwoWay à un TextBox dans la fenêtre Configuration.
        /// </summary>
        public int AdresseFixe { get; set; }

        public override string ToString() => Libelle;
    }
}
