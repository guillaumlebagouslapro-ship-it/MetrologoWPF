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
}
