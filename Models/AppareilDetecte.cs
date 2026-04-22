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

        /// <summary>Type connu reconnu depuis l'IDN (Stanford/Racal/EIP), ou null si inconnu.</summary>
        public TypeAppareilIEEE? TypeReconnu { get; init; }

        /// <summary>
        /// Modèle du catalogue local correspondant (null si l'appareil n'est pas encore enregistré).
        /// Permet d'afficher "reconnu" dans le diagnostic et de proposer l'enregistrement si null.
        /// </summary>
        public ModeleAppareil? ModeleReconnu { get; set; }

        /// <summary>Vrai si l'appareil est reconnu soit par le catalogue, soit par le legacy hardcodé.</summary>
        public bool EstReconnu => ModeleReconnu != null || TypeReconnu.HasValue;

        public string AdresseCourte => $"GPIB{Board}::{Adresse}";

        public string Libelle => ModeleReconnu?.Nom
            ?? (Modele is null ? $"Inconnu ({AdresseCourte})" : $"{Modele} ({AdresseCourte})");

        /// <summary>
        /// Déduit le type IEEE supporté par le catalogue à partir de l'IDN.
        /// Retourne null si l'appareil n'est pas dans les 3 types historiques.
        /// </summary>
        public static TypeAppareilIEEE? DeviquerType(string? fabricant, string? modele)
        {
            string fab = (fabricant ?? string.Empty).ToUpperInvariant();
            string mod = (modele ?? string.Empty).ToUpperInvariant();

            if (fab.Contains("STANFORD") && mod.Contains("SR620")) return TypeAppareilIEEE.Stanford;
            if (fab.Contains("RACAL") && mod.Contains("1996"))     return TypeAppareilIEEE.Racal;
            if (mod.Contains("545"))                               return TypeAppareilIEEE.EIP;

            return null;
        }
    }
}
