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

    public enum ModeMetrologo
    {
        Exploitation,
        Simulation,
        Validation
    }

    /// <summary>La voie concernée par la mesure. On n'envoie à l'appareil que les réglages de
    /// cette voie-là (impédance, couplage, filtre, trigger).</summary>
    public enum VoieActive
    {
        A,
        B,
        C
    }
}
