using System;
using System.Collections.Generic;

namespace Metrologo.Models
{
    public class Mesure
    {
        public string NumFI { get; set; } = string.Empty;

        /// <summary>Id du modèle catalogue sélectionné, source unique de l'appareil utilisé.
        /// Obligatoire au lancement de la mesure.</summary>
        public string IdModeleCatalogue { get; set; } = string.Empty;

        /// <summary>Adresse GPIB forcée pour le mode adresses fixes (legacy sans *IDN? : EIP/Racal/Stanford).
        /// Si >= 0, l'orchestrator prend cette adresse au lieu du match IDN. -1 = résolution IDN normale.</summary>
        public int AdresseFixeForcee { get; set; } = -1;

        /// <summary>Commandes SCPI des réglages choisis dans la fenêtre Configuration,
        /// rejouées par l'orchestrator après le *RST avant la boucle de mesures.</summary>
        public List<string> CommandesScpiReglages { get; set; } = new();

        /// <summary>Voie active pour la mesure. Les réglages des autres voies ne sont pas envoyés à l'appareil.</summary>
        public VoieActive VoieActive { get; set; } = VoieActive.A;
        public TypeMesure TypeMesure { get; set; }
        public ModeMesure ModeMesure { get; set; }
        public SourceMesure SourceMesure { get; set; } = SourceMesure.Frequencemetre;
        public int NbMesures { get; set; } = 30;

        // Paramètres issus des fichiers Delphi et .ini
        public bool InitManu { get; set; }
        public int InputIndex { get; set; }    // Gamme/Entrée
        public int CouplingIndex { get; set; } // Couplage AC/DC

        // --- Intervalle de temps piloté logiciel (53230A) ---
        // Paramètres SCPI envoyés quand l'init manuelle est décochée. Cf. CONF:TINT.

        /// <summary>false = mesure d'intervalle sur 1 voie (start ET stop sur la voie 1) ;
        /// true = 2 voies (start voie 1, stop voie 2).</summary>
        public bool IntervDeuxVoies { get; set; }

        // Voie 1 (et unique en mode 1 voie)
        public bool IntervDc1 { get; set; } = true;     // true = DC, false = AC
        public bool IntervImp50_1 { get; set; } = true; // true = 50 ohm, false = 1 Mohm
        /// <summary>Seuil de déclenchement du DÉPART (V), porté par la voie 1 (INP1:LEV1).</summary>
        public double IntervSeuilStart { get; set; } = 1.0;
        public bool IntervStartMontant { get; set; } = true; // pente départ : true = montant

        // Stop : 1 voie = INP1:LEV2/SLOP2 ; 2 voies = INP2:LEV1/SLOP1
        /// <summary>Seuil de déclenchement de l'ARRÊT (V).</summary>
        public double IntervSeuilStop { get; set; } = 1.0;
        public bool IntervStopMontant { get; set; } = false; // pente arrêt : défaut descendant (cas largeur 1 voie)

        // Voie 2 (mode 2 voies uniquement)
        public bool IntervDc2 { get; set; } = true;
        public bool IntervImp50_2 { get; set; } = true;

        /// <summary>Hold-off (1 voie uniquement) : inhibe le 1er front pour mesurer jusqu'au
        /// front suivant (cas montant→montant). Exprimé en ns. 0 = désactivé.</summary>
        public double IntervHoldoffNs { get; set; }

        /// <summary>
        /// Temps de porte. Mono-élément pour Fréquence/Intervalle/Tachy ; N gates pour
        /// Stabilité (une feuille Excel par gate, remplace les anciennes procédures auto).
        /// Défaut index 6 (1 s), compat flux Fréquence.
        /// </summary>
        public List<int> GateIndices { get; set; } = new() { 6 };

        /// <summary>Raccourci vers la première gate (mono-gate : Fréquence, Intervalle,
        /// ZNGate/ZNValGateSecondes). L'affectation remplace toute la liste.</summary>
        public int GateIndex
        {
            get => GateIndices.Count > 0 ? GateIndices[0] : 6;
            set => GateIndices = new List<int> { value };
        }

        // Paramètres pour le mode Indirect
        public double FNominale { get; set; } = 10000000;
        public int IndexMultiplicateur { get; set; }

        // Paramètres d'incertitude (saisis post-mesure via SaisiePostMesureFreqWindow)
        public double Resolution { get; set; } = 0.01;
        public double IncertSupp { get; set; } = 0.0;

        // Voie du multiplexeur (1..4). 0 = pas de multiplexeur.
        public int VoieMux { get; set; } = 0;

        /// <summary>
        /// Module d'incertitude (nom CSV sans extension dans Incertitudes\).
        /// Vide = coefficients par défaut. Donne ZNCoeffA/B pour Fréquence/Stab,
        /// ZNCoeffC/D (RPM, formule I29) pour les tachys (A/B viennent de NumModuleIncertitudeFreq).
        /// </summary>
        public string NumModuleIncertitude { get; set; } = string.Empty;

        /// <summary>Module d'incertitude Fréquence auxiliaire (tachymètres uniquement) :
        /// fournit les coeffs A/B (Hz) du fréquencemètre, combinés avec les C/D (RPM)
        /// de NumModuleIncertitude dans le rapport Excel. Ignoré pour les autres types.</summary>
        public string NumModuleIncertitudeFreq { get; set; } = string.Empty;

        public Mesure()
        {
            TypeMesure = TypeMesure.Frequence;
            ModeMesure = ModeMesure.Direct;
        }
    }
}