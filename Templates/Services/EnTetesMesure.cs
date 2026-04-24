using Metrologo.Models;

namespace Metrologo.Services
{
    /// <summary>
    /// En-têtes et libellés affichés dans le fichier Excel ModFeuille selon le type de mesure.
    /// La feuille Récap. n'est JAMAIS modifiée.
    /// </summary>
    public class EnTetesMesure
    {
        public string Unite { get; set; } = "Hz";
        public string EnteteHeure { get; set; } = "HEURE";
        public string EnteteMesuree { get; set; } = "Fréq. Mesurée (Hz)";
        public string EnteteReelle { get; set; } = "Fréq. Réelle (Hz)";
        public string EnteteDelta { get; set; } = "F(i) - F(i+1)";

        public string LabelMoyenne { get; set; } = "Valeur Moy. Réel. (Hz) =";
        public string LabelFreqRef { get; set; } = "Valeur de réf. (Hz) =";
        public string LabelFreqCorr { get; set; } = "Valeur corrigée (Hz) =";
        public string LabelIncertResol { get; set; } = "Incert. de résolution (Hz) =";
        public string LabelIncertGlob { get; set; } = "Incert. Globale (Hz) =";
    }

    public static class EnTetesMesureHelper
    {
        public static EnTetesMesure Pour(TypeMesure type) => type switch
        {
            TypeMesure.Interval => new EnTetesMesure
            {
                Unite = "s",
                EnteteMesuree = "Intervalle mesuré (s)",
                EnteteReelle = "Intervalle réel (s)",
                EnteteDelta = "T(i) - T(i+1)",
                LabelMoyenne = "Interv. Moy. Réel. (s) =",
                LabelFreqRef = "Valeur de réf. (s) =",
                LabelFreqCorr = "Valeur corrigée (s) =",
                LabelIncertResol = "Incert. résol. (s) =",
                LabelIncertGlob = "Incert. Globale (s) ="
            },

            TypeMesure.TachyContact => new EnTetesMesure
            {
                Unite = "tr/min",
                EnteteMesuree = "Vitesse mesurée (tr/min)",
                EnteteReelle = "Vitesse réelle (tr/min)",
                EnteteDelta = "N(i) - N(i+1)",
                LabelMoyenne = "Vit. Moy. Réel. (tr/min) =",
                LabelFreqRef = "Valeur de réf. (tr/min) =",
                LabelFreqCorr = "Valeur corrigée (tr/min) =",
                LabelIncertResol = "Incert. résol. (tr/min) =",
                LabelIncertGlob = "Incert. Globale (tr/min) ="
            },

            TypeMesure.Stroboscope => new EnTetesMesure
            {
                Unite = "Hz",
                EnteteMesuree = "Fréq. stroboscope (Hz)",
                EnteteReelle = "Fréq. réelle (Hz)",
                EnteteDelta = "F(i) - F(i+1)",
                LabelMoyenne = "Fréq. Moy. Réel. (Hz) =",
                LabelFreqRef = "Valeur fréq. réf. (Hz) =",
                LabelFreqCorr = "Valeur fréq. corrigée (Hz) =",
                LabelIncertResol = "Incert. de résolution (Hz) =",
                LabelIncertGlob = "Incert. Globale (Hz) ="
            },

            TypeMesure.Stabilite => new EnTetesMesure
            {
                Unite = "Hz",
                EnteteMesuree = "Fréq. Mesurée (Hz)",
                EnteteReelle = "Fréq. Réelle (Hz)",
                EnteteDelta = "F(i) - F(i+1)",
                LabelMoyenne = "Fréq. Moy. Réel. (Hz) =",
                LabelFreqRef = "Valeur fréq. réf. (Hz) =",
                LabelFreqCorr = "Valeur fréq. corrigée (Hz) =",
                LabelIncertResol = "Incert. de résolution (Hz) =",
                LabelIncertGlob = "Incert. Globale (Hz) ="
            },

            // Fréquence, Fréquence avant interv, Fréquence finale → défaut
            _ => new EnTetesMesure()
        };

        public static string LibelleType(TypeMesure t) => t switch
        {
            TypeMesure.Frequence => "Fréquence",
            TypeMesure.Stabilite => "Stabilité",
            TypeMesure.Interval => "Intervalle de temps",
            TypeMesure.FreqAvantInterv => "Fréquence avant intervention",
            TypeMesure.FreqFinale => "Fréquence finale",
            TypeMesure.TachyContact => "Tachymétrie par contacts",
            TypeMesure.Stroboscope => "Stroboscope",
            _ => t.ToString()
        };

        // Gate time helpers (index → libellé et secondes). L'échelle doit rester alignée avec la
        // combo GateTimes de ConfigurationViewModel/SelectionGateViewModel et avec CatalogueAdapter.
        private static readonly string[] _libellesGate =
        {
            "10 ms", "20 ms", "50 ms", "100 ms", "200 ms", "500 ms",
            "1 s", "2 s", "5 s", "10 s", "20 s", "50 s",
            "100 s", "200 s", "500 s", "1000 s"
        };

        private static readonly double[] _secondesGate =
        {
            0.010, 0.020, 0.050, 0.100, 0.200, 0.500,
            1.0, 2.0, 5.0, 10.0, 20.0, 50.0,
            100.0, 200.0, 500.0, 1000.0
        };

        public static string LibelleGate(int index) => index switch
        {
            -2 => "Procédure auto (10 ms → 10 s)",
            -1 => "Procédure auto (10 ms → 100 s)",
            >= 0 and var i when i < _libellesGate.Length => _libellesGate[i],
            _ => ""
        };

        public static double SecondesGate(int index) =>
            (index >= 0 && index < _secondesGate.Length) ? _secondesGate[index] : 1.0;
    }
}
