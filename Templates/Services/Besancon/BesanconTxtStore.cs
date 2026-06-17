using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Metrologo.Services.Besancon
{
    /// <summary>Stockage cumulatif des valeurs Besançon (remplace BASE_E2M SQL). Fichier texte
    /// tab-séparé (MJD, date, valeur Hz), trié par MJD, sans doublon.</summary>
    public static class BesanconTxtStore
    {
        /// <summary>Fichier texte cumulatif des valeurs journalières Besançon (partage).</summary>
        public static string CheminValeurs =>
            Path.Combine(CheminsMetrologo.Besancon, "valeurs_besancon.txt");

        private const string EnTete = "MJD\tDate\tValeur(Hz)";

        /// <summary>Une correction détectée : la valeur d'un MJD déjà enregistré a changé à la source.</summary>
        public readonly struct Correction
        {
            public int Mjd { get; init; }
            public double Ancienne { get; init; }
            public double Nouvelle { get; init; }
        }

        /// <summary>Bilan d'un AjouterAsync : nouvelles dates + corrections appliquées.</summary>
        public readonly struct ResultatAjout
        {
            public int Nouvelles { get; init; }
            public IReadOnlyList<Correction> Corrections { get; init; }
            public bool FichierModifie => Nouvelles > 0 || (Corrections?.Count ?? 0) > 0;
        }

        /// <summary>Fusionne les mesures FTP dans le cumulatif : nouveau MJD = ajout,
        /// valeur différente = correction (Besançon révise parfois), identique = ignoré.
        /// Réécriture uniquement si le fichier est modifié.</summary>
        public static async Task<ResultatAjout> AjouterAsync(IEnumerable<MesureBesancon> mesures)
        {
            Directory.CreateDirectory(CheminsMetrologo.Besancon);

            // Charge l'historique pour dédoublonner.
            var parMjd = await LireAsync();

            int nouvelles = 0;
            var corrections = new List<Correction>();
            foreach (var m in mesures)
            {
                if (parMjd.TryGetValue(m.Mjd, out double existante))
                {
                    // Date déjà connue : réécriture seulement si la source a corrigé la valeur.
                    if (!MemeValeur(existante, m.Valeur))
                    {
                        corrections.Add(new Correction { Mjd = m.Mjd, Ancienne = existante, Nouvelle = m.Valeur });
                        parMjd[m.Mjd] = m.Valeur;   // applique la correction
                    }
                    continue;
                }
                parMjd[m.Mjd] = m.Valeur;   // nouvelle date
                nouvelles++;
            }

            var res = new ResultatAjout { Nouvelles = nouvelles, Corrections = corrections };
            if (!res.FichierModifie) return res;   // rien de neuf ni corrigé : pas de réécriture

            var sb = new StringBuilder();
            sb.AppendLine(EnTete);
            foreach (var kv in parMjd)   // SortedDictionary : trié par MJD
                sb.AppendLine(FormaterLigne(kv.Key, kv.Value));

            await File.WriteAllTextAsync(CheminValeurs, sb.ToString(), Encoding.UTF8);
            return res;
        }

        /// <summary>Égalité par représentation texte (InvariantCulture) : évite le bruit
        /// flottant tout en détectant les vraies corrections.</summary>
        private static bool MemeValeur(double a, double b) =>
            a.ToString(CultureInfo.InvariantCulture) == b.ToString(CultureInfo.InvariantCulture);

        /// <summary>Relit le fichier cumulatif en dictionnaire MJD -> valeur trié.
        /// Vide si absent ; en-tête et lignes non conformes ignorés.</summary>
        public static async Task<SortedDictionary<int, double>> LireAsync()
        {
            var parMjd = new SortedDictionary<int, double>();
            if (!File.Exists(CheminValeurs)) return parMjd;

            string[] lignes = await File.ReadAllLinesAsync(CheminValeurs);
            foreach (var brut in lignes)
            {
                string ligne = brut.Trim();
                if (ligne.Length == 0) continue;

                var tokens = ligne.Split('\t');
                if (tokens.Length < 3) continue;
                // 1er token entier = MJD (ignore l'en-tête et tout texte parasite).
                if (!int.TryParse(tokens[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out int mjd))
                    continue;
                if (!double.TryParse(tokens[2], NumberStyles.Float, CultureInfo.InvariantCulture, out double v))
                    continue;
                parMjd[mjd] = v;
            }
            return parMjd;
        }

        private static string FormaterLigne(int mjd, double valeur) =>
            $"{mjd}\t{JourJulien.DepuisMjd(mjd):dd/MM/yyyy}\t{valeur.ToString(CultureInfo.InvariantCulture)}";
    }
}
