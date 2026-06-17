namespace Metrologo.Models
{
    /// <summary>Un appareil IEEE-488 trouvé lors du scan GPIB. Simple POCO, rangé dans
    /// EtatApplication.AppareilsDetectes et affiché dans la fenêtre de configuration.</summary>
    public class AppareilDetecte
    {
        public int Board { get; init; }
        public int Adresse { get; init; }
        public string Ressource { get; init; } = string.Empty;

        /// <summary>Interface par laquelle l'appareil a répondu (GPIB ou LAN).</summary>
        public TypeBus TypeBus { get; init; } = TypeBus.Gpib;

        /// <summary>Hôte/IP d'un appareil LAN (null en GPIB).</summary>
        public string? Hote { get; init; }

        public string? IdnBrut { get; init; }
        public string? Fabricant { get; init; }
        public string? Modele { get; init; }
        public string? NumeroSerie { get; init; }
        public string? Firmware { get; init; }

        /// <summary>
        /// Le modèle du catalogue local qui correspond (null tant que l'appareil n'est pas enregistré).
        /// Ça permet d'afficher "reconnu" dans le diagnostic, et de proposer l'enregistrement quand c'est null.
        /// </summary>
        public ModeleAppareil? ModeleReconnu { get; set; }

        public bool EstReconnu => ModeleReconnu != null;

        /// <summary>Vrai quand la réponse à *IDN? a l'air incohérente — le signe probable de deux appareils
        /// partageant la même adresse GPIB (collision de bus). C'est le scan qui le renseigne.</summary>
        public bool ConflitAdressePossible { get; init; }

        public string AdresseCourte => TypeBus == TypeBus.Lan
            ? $"LAN {Hote}"
            : $"GPIB{Board}::{Adresse}";

        public string Libelle => ModeleReconnu?.Nom
            ?? (Modele is null ? $"Inconnu ({AdresseCourte})" : $"{Modele} ({AdresseCourte})");
    }
}
