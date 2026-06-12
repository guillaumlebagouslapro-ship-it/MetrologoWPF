using System;
using System.Collections.Generic;
using System.Globalization;

namespace Metrologo.Services.Besancon
{
    /// <summary>Une valeur du fichier de Besançon : date julienne modifiée + valeur corrigée.</summary>
    public sealed class MesureBesancon
    {
        public int Mjd { get; set; }
        public double Valeur { get; set; }
    }

    /// <summary>
    /// Parse le fichier ef_utcop récupéré sur le FTP de l'observatoire de Besançon. Format hérité
    /// du legacy Delphi : une ligne par jour, date julienne puis valeur, séparées par des espaces.
    /// La 1re ligne (en-tête) peut être sautée comme le faisait le Delphi (boucle de 1 à Count-1) ;
    /// les lignes non conformes sont ignorées.
    /// </summary>
    public static class BesanconParser
    {
        public static List<MesureBesancon> Parser(string contenu, bool ignorerPremiereLigne = false)
        {
            var resultats = new List<MesureBesancon>();
            if (string.IsNullOrWhiteSpace(contenu)) return resultats;

            var lignes = contenu.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            for (int i = 0; i < lignes.Length; i++)
            {
                if (ignorerPremiereLigne && i == 0) continue;

                string ligne = lignes[i].Trim();
                if (ligne.Length == 0) continue;

                // ef_utcop commence par plusieurs lignes de commentaire préfixées par #
                // (instabilité, sauts de phase) : on les saute
                if (ligne[0] == '#' || ligne[0] == ';') continue;

                var tokens = ligne.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                if (tokens.Length < 2) continue;

                // 1er token : date julienne (on garde la partie entière si décimale)
                string sDate = tokens[0];
                int ptDate = sDate.IndexOf('.');
                if (ptDate > 0) sDate = sDate.Substring(0, ptDate);
                if (!int.TryParse(sDate, NumberStyles.Integer, CultureInfo.InvariantCulture, out int mjd))
                    continue;

                // 2e token : valeur (on tolère ',' ou '.' comme séparateur décimal)
                string sVal = tokens[1].Replace(',', '.');
                if (!double.TryParse(sVal, NumberStyles.Float, CultureInfo.InvariantCulture, out double valeur))
                    continue;

                resultats.Add(new MesureBesancon { Mjd = mjd, Valeur = valeur });
            }
            return resultats;
        }
    }
}
