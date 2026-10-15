using System.IO;
using System.Text.Json;

namespace Metrologo.Services
{
    /// <summary>
    /// Estime la durée d'un balayage de stabilité, pour l'afficher avant le lancement et donner
    /// une heure de fin pendant une série de répétitions. Purement indicatif : les répétitions
    /// s'enchaînent sur la fin RÉELLE de la précédente (await), jamais sur cette estimation.
    /// <para>
    /// Le surcoût de chaque temps de porte (temps passé en plus de nbMesures × porte : dialogue
    /// GPIB, temps mort entre deux mesures) est APPRIS sur les mesures réelles de ce poste : le
    /// mode d'acquisition change selon la porte (lot sous 1 s, une par une à partir de 1 s) et selon
    /// l'appareil, aucune constante ne tombait juste pour tous les temps. Appris dans
    /// <c>Configuration\durees-stab.json</c> ; valeurs par défaut tant que rien n'est mesuré.
    /// </para>
    /// </summary>
    public static class EstimationDureeStab
    {
        /// <summary>Init de l'appareil (*RST, configuration, rejeu SCPI) en début de balayage (défaut).</summary>
        private const double SECONDES_INIT = 8.0;
        /// <summary>Par gate : fermeture/préparation/sauvegarde/réouverture du classeur + programmation.</summary>
        private const double SECONDES_PAR_GATE = 6.0;
        /// <summary>Surcoût par mesure par défaut : lot gap-free (porte < 1 s) / une par une (porte ≥ 1 s).</summary>
        private const double SURCOUT_DEFAUT_LOT = 0.06;
        private const double SURCOUT_DEFAUT_UNITAIRE = 0.4;
        /// <summary>Marge : 10 % de la durée calculée, plafonnée à 10 min (l'ancienne marge fixe de
        /// 10 min par stab écrasait l'estimation dès qu'on retirait des temps de porte).</summary>
        private const double PART_MARGE = 0.10;
        private const double SECONDES_MARGE_MAX = 10 * 60;
        /// <summary>Poids d'une nouvelle mesure dans la moyenne glissante apprise.</summary>
        private const double POIDS_NOUVELLE_MESURE = 0.5;

        private static readonly object _verrou = new();
        private static DureesApprises? _appris;

        private static string Fichier => Path.Combine(CheminsMetrologo.Configuration, "durees-stab.json");

        /// <summary>Durée estimée d'UN balayage, marge comprise : init + chaque gate (classeur +
        /// nbMesures × (porte + surcoût appris de cette porte)) + 10 % (≤ 10 min).</summary>
        public static TimeSpan EstimerBalayage(IEnumerable<int> gateIndices, int nbMesures)
        {
            var gates = gateIndices?.ToList() ?? new List<int>();
            if (gates.Count == 0 || nbMesures <= 0) return TimeSpan.Zero;

            double secondes;
            lock (_verrou)
            {
                var appris = Charger();
                secondes = appris.SecondesInit ?? SECONDES_INIT;
                foreach (int g in gates)
                    secondes += SECONDES_PAR_GATE
                              + nbMesures * (EnTetesMesureHelper.SecondesGate(g) + Surcout(appris, g));
            }
            secondes += Math.Min(SECONDES_MARGE_MAX, secondes * PART_MARGE);
            return TimeSpan.FromSeconds(secondes);
        }

        /// <summary>Appelé par l'orchestrateur à la fin RÉUSSIE d'une gate de stab : déduit le
        /// surcoût réel par mesure de cette porte et l'intègre à la moyenne apprise.</summary>
        public static void EnregistrerGate(int gateIndex, int nbMesures, TimeSpan duree)
        {
            if (nbMesures <= 0) return;
            double surcout = (duree.TotalSeconds - SECONDES_PAR_GATE) / nbMesures
                           - EnTetesMesureHelper.SecondesGate(gateIndex);
            surcout = Math.Clamp(surcout, 0, 10);
            lock (_verrou)
            {
                var appris = Charger();
                appris.SurcoutParMesure[gateIndex.ToString()] = appris.SurcoutParMesure.TryGetValue(gateIndex.ToString(), out var ancien)
                    ? ancien + POIDS_NOUVELLE_MESURE * (surcout - ancien)
                    : surcout;
                Sauver(appris);
            }
            // Hors verrou : l'abonné (SuiviFinStab) rappelle EstimerBalayage sous son propre verrou.
            try { GateEnregistree?.Invoke(gateIndex); } catch { /* indicatif, ne doit rien casser */ }
        }

        /// <summary>Une gate de stab vient de se terminer (sert à recaler l'heure de fin affichée).</summary>
        public static event Action<int>? GateEnregistree;

        /// <summary>Appelé au début de la 1re gate : durée réelle de l'init de l'appareil.</summary>
        public static void EnregistrerInit(TimeSpan duree)
        {
            double s = Math.Clamp(duree.TotalSeconds, 0, 120);
            lock (_verrou)
            {
                var appris = Charger();
                appris.SecondesInit = appris.SecondesInit is double ancien
                    ? ancien + POIDS_NOUVELLE_MESURE * (s - ancien)
                    : s;
                Sauver(appris);
            }
        }

        private static double Surcout(DureesApprises appris, int gateIndex) =>
            appris.SurcoutParMesure.TryGetValue(gateIndex.ToString(), out var s)
                ? s
                : EnTetesMesureHelper.SecondesGate(gateIndex) >= 1.0 ? SURCOUT_DEFAUT_UNITAIRE : SURCOUT_DEFAUT_LOT;

        private static DureesApprises Charger()
        {
            if (_appris != null) return _appris;
            try
            {
                if (File.Exists(Fichier))
                    _appris = JsonSerializer.Deserialize<DureesApprises>(File.ReadAllText(Fichier));
            }
            catch { /* fichier illisible : on repart des valeurs par défaut */ }
            return _appris ??= new DureesApprises();
        }

        private static void Sauver(DureesApprises appris)
        {
            try
            {
                Directory.CreateDirectory(CheminsMetrologo.Configuration);
                File.WriteAllText(Fichier,
                    JsonSerializer.Serialize(appris, new JsonSerializerOptions { WriteIndented = true }));
            }
            catch { /* best-effort : l'estimation reste indicative */ }
        }

        /// <summary>Contenu de durees-stab.json : surcoût par mesure (s) par index de porte, init (s).</summary>
        private sealed class DureesApprises
        {
            public Dictionary<string, double> SurcoutParMesure { get; set; } = new();
            public double? SecondesInit { get; set; }
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
