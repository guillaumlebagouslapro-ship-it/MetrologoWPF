namespace Metrologo.Services.Incertitude
{
    /// <summary>
    /// Une ligne du tableau d'incertitude d'un module : pour une combinaison Fonction +
    /// temps de mesure + plage de fréquence, on a les 2 coefficients utilisés par le
    /// calcul d'incertitude globale (ZNCoeffA = relative, ZNCoeffB = absolue).
    ///
    /// Reflet exact d'une ligne du tableau papier "Feuille Recap Incertitude CEAO du
    /// module XXX" fourni par les métrologues.
    /// </summary>
    public class LigneIncertitude
    {
        /// <summary>Type de mesure auquel s'applique cette ligne. Convention :
        /// "Freq", "FreqAv", "FreqFin", "Stab", "Interv", "TachyC", "Strobo".</summary>
        public string Fonction { get; set; } = string.Empty;

        /// <summary>Temps de mesure en secondes (gate) pour lequel cette ligne est valide.
        /// = "Condition optionnelle 1" du tableau C.E.A.O.</summary>
        public double TempsDeMesure { get; set; }

        /// <summary>Condition optionnelle 2 — texte libre (ex. "CONNEXION", "AC/DC", etc.).
        /// Vide pour les modules qui n'utilisent pas cette dimension.</summary>
        public string Condition2 { get; set; } = string.Empty;

        /// <summary>Borne basse (incluse) du domaine de mesure 1 (= Fréquence par défaut), en Hz.</summary>
        public double BorneBasse { get; set; }

        /// <summary>Borne haute (incluse) du domaine de mesure 1, en Hz.</summary>
        public double BorneHaute { get; set; }

        /// <summary>Borne basse du domaine de mesure 2 (optionnel — 0 si non utilisé).
        /// Permet de modéliser un module qui couvre 2 plages parallèles avec les mêmes coeffs.</summary>
        public double BorneBasseDomaine2 { get; set; }

        /// <summary>Borne haute du domaine 2 (0 si non utilisé).</summary>
        public double BorneHauteDomaine2 { get; set; }

        /// <summary>Incertitude relative (sans dimension) — devient <c>ZNCoeffA</c> dans le Excel.</summary>
        public double IncertRelative { get; set; }

        /// <summary>Incertitude absolue (en Hz) — devient <c>ZNCoeffB</c> dans le Excel.</summary>
        public double IncertAbsolue { get; set; }

        /// <summary>
        /// Vrai si une fréquence donnée tombe dans le domaine 1 ou dans le domaine 2
        /// (si renseigné). Le domaine 2 est ignoré quand BorneHauteDomaine2 = 0.
        /// </summary>
        public bool Couvre(double frequenceHz)
        {
            if (frequenceHz >= BorneBasse && frequenceHz <= BorneHaute) return true;
            if (BorneHauteDomaine2 > 0 &&
                frequenceHz >= BorneBasseDomaine2 && frequenceHz <= BorneHauteDomaine2) return true;
            return false;
        }
    }
}
