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

        /// <summary>
        /// Temps de porte de la mesure. Un seul élément pour Fréquence/Intervalle/Tachy ; pour la
        /// Stabilité, N gates balayées séquentiellement (une feuille Excel par gate), ce qui remplace
        /// les anciennes procédures auto. Défaut index 6 (1 s), compat avec le flux Fréquence existant.
        /// </summary>
        public List<int> GateIndices { get; set; } = new() { 6 };

        /// <summary>
        /// Raccourci vers la première gate, pour tous les usages mono-gate (Fréquence, Intervalle,
        /// zones nommées ZNGate/ZNValGateSecondes). L'écriture remplace toute la liste par un élément.
        /// </summary>
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
        /// Module d'incertitude sélectionné (nom du CSV sans extension dans
        /// %LocalAppData%\Metrologo\Incertitudes\). Vide = coefficients hardcoded par défaut.
        /// Fournit ZNCoeffA/B pour Fréquence/Stab, ou ZNCoeffC/D (côté RPM, formule I29) pour les
        /// tachymètres, dont les coeffs A/B viennent alors de NumModuleIncertitudeFreq.
        /// </summary>
        public string NumModuleIncertitude { get; set; } = string.Empty;

        /// <summary>
        /// Module d'incertitude Fréquence auxiliaire, uniquement pour les tachymètres : donne
        /// l'incertitude A/B (en Hz) du fréquencemètre sous-jacent, combinée dans le rapport Excel
        /// avec les C/D (RPM) de NumModuleIncertitude. Ignoré pour les autres types.
        /// </summary>
        public string NumModuleIncertitudeFreq { get; set; } = string.Empty;

        public Mesure()
        {
            TypeMesure = TypeMesure.Frequence;
            ModeMesure = ModeMesure.Direct;
        }
    }
}