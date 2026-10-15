using Metrologo.Models;

namespace Metrologo.Services.Incertitude
{
    /// <summary>
    /// Correspondance <see cref="TypeMesure"/> ↔ libellé « Fonction » des CSV de modules d'incertitude
    /// (mapping unique, partagé par l'UI de Configuration et le runtime ExcelService).
    /// </summary>
    public static class IncertitudeFonctionHelper
    {
        /// <summary>
        /// Libellé court de la colonne <c>Fonction</c> des CSV, à garder aligné sur
        /// <c>FonctionsDisponibles</c> de <see cref="ViewModels.GestionModulesIncertitudeViewModel"/>.
        /// </summary>
        public static string NomFonction(TypeMesure type) => type switch
        {
            TypeMesure.Frequence       => "Freq",
            TypeMesure.FreqAvantInterv => "FreqAv",
            TypeMesure.FreqFinale      => "FreqFin",
            TypeMesure.Stabilite       => "Stab",
            TypeMesure.Interval        => "Interv",
            TypeMesure.TachyContact    => "TachyC",
            TypeMesure.TachyOptique    => "TachyO",
            TypeMesure.Stroboscope     => "Strobo",
            _                          => string.Empty
        };
    }
}
