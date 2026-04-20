using System;

namespace Metrologo.Models
{
    public class Mesure
    {
        public string NumFI { get; set; } = string.Empty;
        public TypeAppareilIEEE Frequencemetre { get; set; }
        public TypeMesure TypeMesure { get; set; }
        public ModeMesure ModeMesure { get; set; }
        public SourceMesure SourceMesure { get; set; } = SourceMesure.Frequencemetre;
        // Dans Models/Mesure.cs, ajoutez cette ligne dans la classe :
        public int NbMesures { get; set; } = 30;

        // Paramètres issus des fichiers Delphi et .ini
        public bool InitManu { get; set; }
        public int InputIndex { get; set; }    // Gamme/Entrée
        public int CouplingIndex { get; set; } // Couplage AC/DC
                                               // Remplacez la ligne existante par celle-ci :
        public int GateIndex { get; set; } = 6; // Index 6 correspond à "1 s" dans la nouvelle liste complète

        // Paramètres pour le mode Indirect
        public double FNominale { get; set; } = 10000000;
        public int IndexMultiplicateur { get; set; }

        // Paramètres d'incertitude (ajustables via dialogue ParamsIncert)
        public double Resolution { get; set; } = 0.01;
        public double IncertSupp { get; set; } = 0.0;

        public Mesure()
        {
            TypeMesure = TypeMesure.Frequence;
            Frequencemetre = TypeAppareilIEEE.Stanford;
            ModeMesure = ModeMesure.Direct;
        }
    }
}