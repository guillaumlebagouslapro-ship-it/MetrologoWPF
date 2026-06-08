using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Metrologo.Services.Besancon
{
    /// <summary>
    /// Stockage cumulatif des valeurs de Besançon dans un simple fichier texte sur le partage,
    /// en remplacement de la base SQL <c>BASE_E2M</c>. Une ligne par mesure
    /// (<c>MJD &lt;tab&gt; date &lt;tab&gt; valeur(Hz)</c>), triée par date julienne croissante et
    /// SANS doublon de MJD : à chaque récupération FTP, seules les dates encore absentes sont
    /// ajoutées — le fichier s'incrémente donc jour après jour.
    /// </summary>
    public static class BesanconTxtStore
    {
        /// <summary>Fichier texte cumulatif des valeurs journalières Besançon (partage).</summary>
        public static string CheminValeurs =>
            Path.Combine(CheminsMetrologo.Besancon, "valeurs_besancon.txt");

        private const string EnTete = "MJD\tDate\tValeur(Hz)";

        /// <summary>
        /// Ajoute au fichier cumulatif les mesures dont le MJD n'y figure pas déjà. Retourne le
        /// nombre de NOUVELLES lignes écrites. Le fichier est réécrit en entier, trié par MJD
        /// croissant et précédé d'un en-tête — ce qui garantit l'absence de doublon même si le
        /// fichier FTP renvoie plusieurs fois les mêmes dates.
        /// </summary>
        public static async Task<int> AjouterAsync(IEnumerable<MesureBesancon> mesures)
        {
            Directory.CreateDirectory(CheminsMetrologo.Besancon);

            // Charge l'historique existant (MJD -> valeur) pour dédoublonner et le conserver.
            var parMjd = await LireAsync();

            int nouvelles = 0;
            foreach (var m in mesures)
            {
                if (parMjd.ContainsKey(m.Mjd)) continue;   // date déjà enregistrée → ignorée
                parMjd[m.Mjd] = m.Valeur;
                nouvelles++;
            }
            if (nouvelles == 0) return 0;   // rien de neuf → on ne réécrit pas le fichier

            var sb = new StringBuilder();
            sb.AppendLine(EnTete);
            foreach (var kv in parMjd)   // SortedDictionary → déjà trié par MJD croissant
                sb.AppendLine(FormaterLigne(kv.Key, kv.Value));

            await File.WriteAllTextAsync(CheminValeurs, sb.ToString(), Encoding.UTF8);
            return nouvelles;
        }

        /// <summary>
        /// Relit le fichier cumulatif en dictionnaire <c>MJD -&gt; valeur</c> (trié, vide si le
        /// fichier n'existe pas encore). L'en-tête et les lignes non conformes sont ignorés.
        /// </summary>
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
                // 1er token entier = MJD (ignore l'en-tête « MJD » et tout texte parasite).
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
