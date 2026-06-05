using System;
using System.Collections.Generic;

namespace Metrologo.Models
{
    public class Mesure
    {
        public string NumFI { get; set; } = string.Empty;

        /// <summary>
        /// Id du modèle catalogue sélectionné — source unique de l'appareil utilisé pour la mesure.
        /// Vide uniquement lors de la construction initiale, obligatoire au moment de lancer la mesure.
        /// </summary>
        public string IdModeleCatalogue { get; set; } = string.Empty;

        /// <summary>
        /// Adresse GPIB forcée pour le mode « adresses fixes » (appareils legacy sans <c>*IDN?</c> :
        /// EIP/Racal/Stanford). Quand &gt;= 0, l'orchestrator résout l'appareil à cette adresse au
        /// lieu de chercher un device détecté par IDN sur le bus. <c>-1</c> = résolution IDN normale.
        /// </summary>
        public int AdresseFixeForcee { get; set; } = -1;

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

        /// <summary>
        /// Liste des temps de porte à utiliser pour la mesure. Pour les types non-Stabilité
        /// (Fréquence, Intervalle, Tachy…) : un unique élément. Pour la Stabilité, peut
        /// contenir N gates : l'orchestrator créera une feuille Excel par gate et balaiera
        /// la liste séquentiellement (équivalent moderne et générique des « procédures auto »
        /// historiques).
        ///
        /// Index 6 (= « 1 s ») par défaut, conservé pour compat avec le flux Fréquence existant.
        /// </summary>
        public List<int> GateIndices { get; set; } = new() { 6 };

        /// <summary>
        /// Accès court-circuit à la première (et souvent unique) gate. Utilisé partout où une
        /// seule gate est attendue (Fréquence, Intervalle, init Excel d'une feuille donnée,
        /// zones nommées <c>ZNGate</c>/<c>ZNValGateSecondes</c>). En lecture, retourne le
        /// premier indice. En écriture, remplace la liste par un unique élément — les usages
        /// historiques continuent donc à fonctionner sans changement.
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
        /// Identifiant du module d'incertitude sélectionné (= nom du fichier CSV sans extension
        /// dans <c>%LocalAppData%\Metrologo\Incertitudes\</c>). Vide = aucun module sélectionné,
        /// l'<c>ExcelService</c> retombera alors sur les coefficients hardcoded par défaut.
        ///
        /// Pour les types Fréquence/Stab/etc., ce module fournit ZNCoeffA et ZNCoeffB.
        /// Pour les types Tachymètre (Contact/Optique), ce module fournit ZNCoeffC et ZNCoeffD
        /// (côté RPM, formule I29) — les coeffs A/B (côté Hz) viennent du module Fréquence
        /// auxiliaire <see cref="NumModuleIncertitudeFreq"/>.
        /// </summary>
        public string NumModuleIncertitude { get; set; } = string.Empty;

        /// <summary>
        /// Module d'incertitude Fréquence auxiliaire — utilisé uniquement pour les mesures
        /// Tachymètre. Le tachymètre s'appuie sur un fréquencemètre dont l'incertitude (A/B
        /// en Hz) est caractérisée par un module CSV de la catégorie Fréquence ; cette
        /// information est ensuite combinée avec l'incertitude propre au tachymètre (C/D
        /// en RPM, issus de <see cref="NumModuleIncertitude"/>) dans le rapport Excel.
        /// Ignoré pour tous les autres types de mesure.
        /// </summary>
        public string NumModuleIncertitudeFreq { get; set; } = string.Empty;

        public Mesure()
        {
            TypeMesure = TypeMesure.Frequence;
            ModeMesure = ModeMesure.Direct;
        }
    }
}