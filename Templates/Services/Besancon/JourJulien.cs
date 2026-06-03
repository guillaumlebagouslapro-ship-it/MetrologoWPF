using System;

namespace Metrologo.Services.Besancon
{
    /// <summary>
    /// Conversion date ⇄ Modified Julian Date (MJD), comme le legacy Delphi
    /// (<c>DateTimeToModifiedJulianDate</c>). Le fichier de Besançon (<c>ef_utcop</c>)
    /// indexe ses valeurs par date julienne modifiée (entier).
    ///
    /// MJD = nombre de jours depuis le 17/11/1858 00:00 UTC.
    /// </summary>
    public static class JourJulien
    {
        private static readonly DateTime EpochMjd = new DateTime(1858, 11, 17);

        /// <summary>Date → MJD (entier, partie jour).</summary>
        public static int VersMjd(DateTime date) => (int)(date.Date - EpochMjd).TotalDays;

        /// <summary>MJD → date (minuit).</summary>
        public static DateTime DepuisMjd(int mjd) => EpochMjd.AddDays(mjd);
    }
}
