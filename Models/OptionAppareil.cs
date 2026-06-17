namespace Metrologo.Models
{
    /// <summary>
    /// Une entrée du menu déroulant "Fréquencemètre" de la fenêtre Configuration.
    /// Elle correspond à un appareil vu sur le bus GPIB, qu'on rattache éventuellement à un modèle du catalogue.
    /// </summary>
    public class OptionAppareil
    {
        /// <summary>Libellé affiché dans le menu déroulant.</summary>
        public string Libelle { get; init; } = string.Empty;

        /// <summary>Les infos de détection sur le bus. Null si l'appareil n'a jamais été détecté.</summary>
        public AppareilDetecte? Detecte { get; init; }

        public bool EstDetecte => Detecte != null;
        public bool EstConnu => Detecte?.ModeleReconnu != null;
        public string AdresseAffiche => Detecte?.AdresseCourte ?? "non connecté";

        // ---------------- mode adresses fixes (appareils legacy non détectés) ----------------

        /// <summary>
        /// Le modèle legacy du catalogue rattaché en mode adresses fixes (l'appareil n'est pas sur le
        /// bus, on le pilote à une adresse qu'on a saisie). Null en mode scan.
        /// </summary>
        public ModeleAppareil? ModeleFixe { get; init; }

        public bool EstFixe => ModeleFixe != null;

        /// <summary>
        /// L'adresse GPIB éditable (mode adresses fixes), pré-remplie depuis le profil. En écriture,
        /// car elle est liée en TwoWay à un TextBox de la fenêtre Configuration.
        /// </summary>
        public int AdresseFixe { get; set; }

        public override string ToString() => Libelle;
    }
}
