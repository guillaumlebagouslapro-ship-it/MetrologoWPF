using Metrologo.Models;

namespace Metrologo.Services.Incertitude
{
    /// <summary>
    /// Conversion <see cref="TypeMesure"/> ↔ libellé "Fonction" utilisé dans les CSV de modules
    /// d'incertitude. Centralise le mapping pour éviter qu'il soit dupliqué (Configuration UI,
    /// runtime ExcelService).
    /// </summary>
    public static class IncertitudeFonctionHelper
    {
        /// <summary>
        /// Renvoie le libellé court de fonction utilisé dans les colonnes <c>Fonction</c>
        /// des CSV de modules. Cohérent avec <c>FonctionsDisponibles</c> de la fenêtre admin
        /// <see cref="ViewModels.GestionModulesIncertitudeViewModel"/>.
        /// </summary>
        public static string NomFonction(TypeMesure type) => type switch
        {
            TypeMesure.Frequence       => "Freq",
            TypeMesure.FreqAvantInterv => "FreqAv",
            TypeMesure.FreqFinale      => "FreqFin",
            TypeMesure.Stabilite       => "Stab",
            TypeMesure.Interval        => "Interv",
            TypeMesure.TachyContact    => "TachyC",
            TypeMesure.Stroboscope     => "Strobo",
            _                          => string.Empty
        };
    }
}
