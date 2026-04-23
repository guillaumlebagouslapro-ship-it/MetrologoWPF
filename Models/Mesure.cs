using System;
using System.Collections.Generic;

namespace Metrologo.Models
{
    public class Mesure
    {
        public string NumFI { get; set; } = string.Empty;
        public TypeAppareilIEEE Frequencemetre { get; set; }

        /// <summary>
        /// Id du modèle catalogue sélectionné (prioritaire sur <see cref="Frequencemetre"/> à l'exécution).
        /// Vide = utiliser les 3 types historiques chargés depuis Metrologo.ini.
        /// </summary>
        public string IdModeleCatalogue { get; set; } = string.Empty;

        /// <summary>
        /// Commandes SCPI correspondant aux réglages choisis dans la fenêtre Configuration
        /// (Impédance, Couplage, Filtre, Trigger, Mode). L'orchestrator les rejoue après le
        /// <c>*RST</c> pour que l'appareil retrouve l'état configuré avant la boucle de mesures.
        /// </summary>
        public List<string> CommandesScpiReglages { get; set; } = new();

        /// <summary>
        /// Voie active pour la mesure. Les réglages des autres voies (A/B/C) ne sont pas envoyés
        /// à l'appareil, pour éviter de modifier une voie que l'utilisateur n'utilise pas.
        /// </summary>
        public VoieActive VoieActive { get; set; } = VoieActive.A;
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

        // Voie du multiplexeur (1..4). 0 = pas de multiplexeur.
        public int VoieMux { get; set; } = 0;

        public Mesure()
        {
            TypeMesure = TypeMesure.Frequence;
            Frequencemetre = TypeAppareilIEEE.Stanford;
            ModeMesure = ModeMesure.Direct;
        }
    }
}