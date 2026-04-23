namespace Metrologo.Models
{
    public enum TypeAppareilIEEE
    {
        Stanford,
        Racal,
        EIP
    }

    public enum TypeMesure
    {
        Frequence,
        Stabilite,
        Interval,
        FreqAvantInterv,
        FreqFinale,
        TachyContact,
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

    /// <summary>
    /// Voie active sur laquelle porte la mesure. Seuls les réglages (impédance, couplage, filtre, trigger)
    /// de cette voie seront appliqués à l'appareil — les autres voies restent intactes.
    /// </summary>
    public enum VoieActive
    {
        A,
        B,
        C
    }
}
