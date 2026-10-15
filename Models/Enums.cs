namespace Metrologo.Models
{
    public enum TypeMesure
    {
        Frequence,
        Stabilite,
        Interval,
        FreqAvantInterv,
        FreqFinale,
        TachyContact,
        TachyOptique,
        Stroboscope
    }

    public enum ModeMesure
    {
        Direct,
        Indirect
    }

    public enum SourceMesure
    {
        Frequencemetre,
        Generateur
    }

    /// <summary>Format des rapports Excel en sortie de mesure (réglage Admin, commun aux postes).</summary>
    public enum FormatRapport
    {
        /// <summary>Excel 97-2003 (.xls) : seul format lu par le logiciel externe d'extraction pour l'instant.</summary>
        Xls,
        /// <summary>Excel actuel (.xlsx) : le fichier de travail de Metrologo, gardé tel quel.</summary>
        Xlsx
    }

    /// <summary>La voie concernée par la mesure. On n'envoie à l'appareil que les réglages de
    /// cette voie-là (impédance, couplage, filtre, trigger).</summary>
    public enum VoieActive
    {
        A,
        B,
        C
    }

    /// <summary>Interface physique par laquelle un appareil est joignable.</summary>
    public enum TypeBus
    {
        Gpib,
        Lan
    }
}
