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

    /// <summary>Voie sur laquelle porte la mesure. Seuls les réglages de cette voie
    /// (impédance, couplage, filtre, trigger) sont envoyés à l'appareil.</summary>
    public enum VoieActive
    {
        A,
        B,
        C
    }
}
