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

        /// <summary>Une correction détectée : la valeur d'un MJD déjà enregistré a changé à la source.</summary>
        public readonly struct Correction
        {
            public int Mjd { get; init; }
            public double Ancienne { get; init; }
            public double Nouvelle { get; init; }
        }

        /// <summary>Bilan d'un <see cref="AjouterAsync"/> : nouvelles dates + corrections appliquées.</summary>
        public readonly struct ResultatAjout
        {
            public int Nouvelles { get; init; }
            public IReadOnlyList<Correction> Corrections { get; init; }
            public bool FichierModifie => Nouvelles > 0 || (Corrections?.Count ?? 0) > 0;
        }

        /// <summary>
        /// Fusionne les mesures FTP dans le fichier cumulatif :
        /// <list type="bullet">
        /// <item>MJD absent → nouvelle ligne (incrément normal du jour) ;</item>
        /// <item>MJD déjà présent mais valeur DIFFÉRENTE → correction appliquée (rare : Besançon
        ///       révise parfois une valeur passée) ;</item>
        /// <item>MJD déjà présent, valeur identique → ignoré (cas courant à chaque récupération).</item>
        /// </list>
        /// Le fichier n'est réécrit (trié par MJD croissant, avec en-tête) que s'il y a au moins
        /// une nouveauté ou une correction. Retourne le bilan détaillé.
        /// </summary>
        public static async Task<ResultatAjout> AjouterAsync(IEnumerable<MesureBesancon> mesures)
        {
            Directory.CreateDirectory(CheminsMetrologo.Besancon);

            // Charge l'historique existant (MJD -> valeur) pour dédoublonner et le conserver.
            var parMjd = await LireAsync();

            int nouvelles = 0;
            var corrections = new List<Correction>();
            foreach (var m in mesures)
            {
                if (parMjd.TryGetValue(m.Mjd, out double existante))
                {
                    // Date déjà connue : on ne réécrit QUE si la source a corrigé la valeur.
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
            if (!res.FichierModifie) return res;   // rien de neuf ni corrigé → pas de réécriture

            var sb = new StringBuilder();
            sb.AppendLine(EnTete);
            foreach (var kv in parMjd)   // SortedDictionary → déjà trié par MJD croissant
                sb.AppendLine(FormaterLigne(kv.Key, kv.Value));

            await File.WriteAllTextAsync(CheminValeurs, sb.ToString(), Encoding.UTF8);
            return res;
        }

        /// <summary>
        /// Égalité « telle qu'écrite dans le fichier » : on compare la représentation
        /// round-trippable (InvariantCulture), exactement ce qui serait persisté. Évite tout
        /// faux positif de bruit flottant tout en détectant la moindre correction réelle.
        /// </summary>
        private static bool MemeValeur(double a, double b) =>
            a.ToString(CultureInfo.InvariantCulture) == b.ToString(CultureInfo.InvariantCulture);

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
