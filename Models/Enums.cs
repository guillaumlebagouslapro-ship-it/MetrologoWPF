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
        TachyOptique,
        Stroboscope
    }

    public enum ModeMesure
    {
        Direct,
        Indirect
    }

    public enum ModeMetrologo
    {
        Exploitation,
        Simulation,
        Validation
    }
}