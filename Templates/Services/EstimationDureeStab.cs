using System;
using System.Collections.Generic;
using System.Linq;

namespace Metrologo.Services
{
    /// <summary>
    /// Estime la durée d'un balayage de stabilité, pour l'afficher avant le lancement et donner
    /// une heure de fin pendant une série de répétitions. Purement indicatif : les répétitions
    /// s'enchaînent sur la fin RÉELLE de la précédente (await), jamais sur cette estimation.
    /// </summary>
    public static class EstimationDureeStab
    {
        // Ordres de grandeur, à ajuster d'après le journal (STAB_SERIE_REPETITION_FIN donne la
        // durée réelle de chaque balayage).
        /// <summary>Init de l'appareil (*RST, configuration, rejeu SCPI) en début de balayage.</summary>
        private const double SECONDES_INIT = 8.0;
        /// <summary>Par gate : fermeture/préparation/sauvegarde/réouverture du classeur + programmation.</summary>
        private const double SECONDES_PAR_GATE = 6.0;
        /// <summary>Par mesure : aller-retour GPIB en plus du temps de porte.</summary>
        private const double SECONDES_PAR_MESURE = 0.06;

        /// <summary>Durée estimée d'UN balayage : chaque gate = nbMesures × (temps de porte + GPIB).</summary>
        public static TimeSpan EstimerBalayage(IEnumerable<int> gateIndices, int nbMesures)
        {
            var gates = gateIndices?.ToList() ?? new List<int>();
            if (gates.Count == 0 || nbMesures <= 0) return TimeSpan.Zero;

            double secondes = SECONDES_INIT;
            foreach (int g in gates)
                secondes += SECONDES_PAR_GATE
                          + nbMesures * (EnTetesMesureHelper.SecondesGate(g) + SECONDES_PAR_MESURE);
            return TimeSpan.FromSeconds(secondes);
        }

        /// <summary>Format lisible : « 45 s », « 12 min », « 3 h 05 ».</summary>
        public static string Formater(TimeSpan duree)
        {
            if (duree.TotalSeconds < 60) return $"{Math.Max(1, (int)Math.Round(duree.TotalSeconds))} s";
            if (duree.TotalMinutes < 60) return $"{(int)Math.Round(duree.TotalMinutes)} min";
            int heures = (int)duree.TotalHours;
            int minutes = (int)Math.Round(duree.TotalMinutes - heures * 60);
            if (minutes == 60) { heures++; minutes = 0; }
            return $"{heures} h {minutes:00}";
        }

        /// <summary>Heure de fin lisible : « 14:32 », ou « demain 03:10 » si on passe minuit.</summary>
        public static string HeureFin(TimeSpan resteAFaire)
        {
            var fin = DateTime.Now + resteAFaire;
            int jours = (fin.Date - DateTime.Today).Days;
            string prefixe = jours switch
            {
                0 => "",
                1 => "demain ",
                _ => $"le {fin:dd/MM} "
            };
            return $"{prefixe}{fin:HH:mm}";
        }
    }
}
