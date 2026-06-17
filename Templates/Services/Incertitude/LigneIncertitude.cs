namespace Metrologo.Services.Incertitude
{
    /// <summary>
    /// Ligne du tableau d'incertitude : pour une combinaison Fonction + temps + plage Hz,
    /// donne les deux coefficients (ZNCoeffA = relatif, ZNCoeffB = absolu).
    /// Reflet d'une ligne du tableau papier "Feuille Recap Incertitude CEAO".
    /// </summary>
    public class LigneIncertitude
    {
        /// <summary>Fonction mesure : "Freq", "FreqAv", "FreqFin", "Stab", "Interv", "TachyC", "Strobo".</summary>
        public string Fonction { get; set; } = string.Empty;

        /// <summary>Gate en secondes pour laquelle cette ligne est valide (= Condition 1 du tableau C.E.A.O.).</summary>
        public double TempsDeMesure { get; set; }

        /// <summary>Condition optionnelle 2 — texte libre (ex. "CONNEXION", "AC/DC", etc.).
        /// Vide pour les modules qui n'utilisent pas cette dimension.</summary>
        public string Condition2 { get; set; } = string.Empty;

        /// <summary>Borne basse (incluse) du domaine de mesure 1 (= Fréquence par défaut), en Hz.</summary>
        public double BorneBasse { get; set; }

        /// <summary>Borne haute (incluse) du domaine de mesure 1, en Hz.</summary>
        public double BorneHaute { get; set; }

        /// <summary>Borne basse du domaine 2 (0 si non utilisé) — même coeffs que le domaine 1.</summary>
        public double BorneBasseDomaine2 { get; set; }

        /// <summary>Borne haute du domaine 2 (0 si non utilisé).</summary>
        public double BorneHauteDomaine2 { get; set; }

        /// <summary>
        /// Incertitude relative (sans dimension) → <c>ZNCoeffA</c> dans l'Excel (Freq/Stab/etc.),
        /// ou <c>ZNCoeffC</c> dans le template tachy (I29 = Vitesse_RPM × C + D).
        /// </summary>
        public double IncertRelative { get; set; }

        /// <summary>
        /// Incertitude absolue (Hz, ou tr/min pour tachy) → <c>ZNCoeffB</c> (Freq/Stab/etc.)
        /// ou <c>ZNCoeffD</c> dans le template tachy (cf. <see cref="IncertRelative"/>).
        /// </summary>
        public double IncertAbsolue { get; set; }

        /// <summary>Vrai si la fréquence tombe dans le domaine 1 ou le domaine 2 (ignoré si BorneHauteDomaine2 = 0).</summary>
        public bool Couvre(double frequenceHz)
        {
            if (frequenceHz >= BorneBasse && frequenceHz <= BorneHaute) return true;
            if (BorneHauteDomaine2 > 0 &&
                frequenceHz >= BorneBasseDomaine2 && frequenceHz <= BorneHauteDomaine2) return true;
            return false;
        }
    }
}
