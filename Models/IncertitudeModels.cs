namespace Metrologo.Models
{
    /// <summary>Les coefficients d'incertitude qui s'appliquent à une plage de fréquence, pour un appareil donné.</summary>
    public class IncertitudeFrequence
    {
        public string Plage { get; set; } = string.Empty;
        public string Raccord { get; set; } = string.Empty;
        public double CoeffA { get; set; }
        public double CoeffB { get; set; }
    }

    /// <summary>Le coefficient d'incertitude associé à un temps de porte, pour un appareil donné.</summary>
    public class IncertitudeStabilite
    {
        public string TempsDeMesure { get; set; } = string.Empty;
        public double Valeur { get; set; }
    }

    /// <summary>Coefficients d'incertitude pour les autres types de mesure : intervalle, tachy, stroboscope.</summary>
    public class IncertitudeAutreMesure
    {
        public string Libelle { get; set; } = string.Empty;
        public double CoeffA { get; set; }
        public double CoeffB { get; set; }
    }

    /// <summary>Les réglages d'incertitude qui valent pour tout l'appareil : nombre de mesures accréditées, etc.</summary>
    public class ParametresIncertitudeGlobaux
    {
        public int NbMesAccr { get; set; } = 30;
        public int TempsMesAccrSec { get; set; } = 10;
        public double IncertAccr { get; set; } = 1e-11;
    }
}
