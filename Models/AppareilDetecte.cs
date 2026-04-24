namespace Metrologo.Models
{
    /// <summary>
    /// Appareil IEEE-488 découvert sur le bus GPIB lors d'un scan.
    /// Simple POCO destiné à être stocké dans <see cref="EtatApplication.AppareilsDetectes"/>
    /// et affiché dans la fenêtre de configuration.
    /// </summary>
    public class AppareilDetecte
    {
        public int Board { get; init; }
        public int Adresse { get; init; }
        public string Ressource { get; init; } = string.Empty;
        public string? IdnBrut { get; init; }
        public string? Fabricant { get; init; }
        public string? Modele { get; init; }
        public string? NumeroSerie { get; init; }
        public string? Firmware { get; init; }

        /// <summary>
        /// Modèle du catalogue local correspondant (null si l'appareil n'est pas encore enregistré).
        /// Permet d'afficher "reconnu" dans le diagnostic et de proposer l'enregistrement si null.
        /// </summary>
        public ModeleAppareil? ModeleReconnu { get; set; }

        public bool EstReconnu => ModeleReconnu != null;

        public string AdresseCourte => $"GPIB{Board}::{Adresse}";

        public string Libelle => ModeleReconnu?.Nom
            ?? (Modele is null ? $"Inconnu ({AdresseCourte})" : $"{Modele} ({AdresseCourte})");
    }
}
