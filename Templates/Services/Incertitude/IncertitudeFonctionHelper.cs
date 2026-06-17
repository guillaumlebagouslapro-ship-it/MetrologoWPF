using Metrologo.Models;

namespace Metrologo.Services.Incertitude
{
    /// <summary>
    /// Fait la correspondance <see cref="TypeMesure"/> ↔ libellé "Fonction" tel qu'on l'écrit dans
    /// les CSV des modules d'incertitude. On regroupe ce mapping ici pour ne pas le retrouver
    /// recopié à plusieurs endroits (l'UI de Configuration, le runtime ExcelService).
    /// </summary>
    public static class IncertitudeFonctionHelper
    {
        /// <summary>
        /// Donne le libellé court de fonction qu'on retrouve dans la colonne <c>Fonction</c> des
        /// CSV de modules. Il reste aligné sur <c>FonctionsDisponibles</c> de la fenêtre admin
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
            TypeMesure.TachyOptique    => "TachyO",
            TypeMesure.Stroboscope     => "Strobo",
            _                          => string.Empty
        };
    }
}
