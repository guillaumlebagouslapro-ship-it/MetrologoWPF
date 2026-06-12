using System;

namespace Metrologo.Services.Besancon
{
    /// <summary>
    /// Conversion date / Modified Julian Date (MJD), comme le DateTimeToModifiedJulianDate du
    /// legacy Delphi. Le fichier ef_utcop de Besançon indexe ses valeurs par MJD entier.
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
