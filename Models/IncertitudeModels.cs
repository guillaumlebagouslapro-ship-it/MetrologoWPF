namespace Metrologo.Models
{
    /// <summary>Coefficient d'incertitude pour une plage de fréquence sur un appareil donné.</summary>
    public class IncertitudeFrequence
    {
        public string Plage { get; set; } = string.Empty;
        public string Raccord { get; set; } = string.Empty;
        public double CoeffA { get; set; }
        public double CoeffB { get; set; }
    }

    /// <summary>Coefficient d'incertitude pour un temps de porte sur un appareil donné.</summary>
    public class IncertitudeStabilite
    {
        public string TempsDeMesure { get; set; } = string.Empty;
        public double Valeur { get; set; }
    }

    /// <summary>Coefficient d'incertitude pour une mesure autre (intervalle, tachy, stroboscope).</summary>
    public class IncertitudeAutreMesure
    {
        public string Libelle { get; set; } = string.Empty;
        public double CoeffA { get; set; }
        public double CoeffB { get; set; }
    }

    /// <summary>Paramètres globaux d'incertitude (nombre de mesures accréditées, etc.).</summary>
    public class ParametresIncertitudeGlobaux
    {
        public int NbMesAccr { get; set; } = 30;
        public int TempsMesAccrSec { get; set; } = 10;
        public double IncertAccr { get; set; } = 1e-11;
    }
}
