using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Ieee;
using Metrologo.Services.Incertitude;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.ViewModels
{
    public partial class ConfigurationViewModel : ObservableObject
    {
        [ObservableProperty]
        private Mesure _mesureConfig = new Mesure();

        [ObservableProperty]
        private bool _estSurBaie = true;

        /// <summary>
        /// Réglages dynamiques du modèle sélectionné (ex: Impédance, Couplage, Filtre).
        /// Peuplée depuis le catalogue à chaque changement d'appareil — l'UI génère une
        /// ComboBox par réglage via un ItemsControl.
        /// </summary>
        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesDynamiques { get; } = new();

        // ------- Collections filtrées par voie, exposées au XAML pour un rendu structuré -------

        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesVoieA { get; } = new();
        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesVoieB { get; } = new();
        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesVoieC { get; } = new();
        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesMode  { get; } = new();

        public bool AReglagesDynamiques => ReglagesDynamiques.Count > 0;
        public bool AReglagesVoieA => ReglagesVoieA.Count > 0;
        public bool AReglagesVoieB => ReglagesVoieB.Count > 0;
        public bool AReglagesVoieC => ReglagesVoieC.Count > 0;
        public bool AReglagesMode  => ReglagesMode.Count > 0;

        // ------- Sélection de la voie active (pilote visibilité + filtrage des commandes envoyées) -------

        public bool EstVoieA
        {
            get => MesureConfig.VoieActive == VoieActive.A;
            set { if (value) { MesureConfig.VoieActive = VoieActive.A; NotifierVoiesActives(); } }
        }

        public bool EstVoieB
        {
            get => MesureConfig.VoieActive == VoieActive.B;
            set { if (value) { MesureConfig.VoieActive = VoieActive.B; NotifierVoiesActives(); } }
        }

        public bool EstVoieC
        {
            get => MesureConfig.VoieActive == VoieActive.C;
            set { if (value) { MesureConfig.VoieActive = VoieActive.C; NotifierVoiesActives(); } }
        }

        /// <summary>Visibilité du bloc Voie A dans le XAML (= réglages définis ET voie active).</summary>
        public bool AfficherVoieA => AReglagesVoieA && EstVoieA;
        public bool AfficherVoieB => AReglagesVoieB && EstVoieB;
        public bool AfficherVoieC => AReglagesVoieC && EstVoieC;

        private void NotifierVoiesActives()
        {
            OnPropertyChanged(nameof(EstVoieA));
            OnPropertyChanged(nameof(EstVoieB));
            OnPropertyChanged(nameof(EstVoieC));
            OnPropertyChanged(nameof(AfficherVoieA));
            OnPropertyChanged(nameof(AfficherVoieB));
            OnPropertyChanged(nameof(AfficherVoieC));
        }

        public ConfigurationViewModel()
        {
            EtatApplication.AppareilsDetectesChange += (_, _) => RebuildAppareils();
            CatalogueAppareilsService.Instance.CatalogueChange += (_, _) => RebuildReglagesDynamiques();
            RebuildAppareils();
            RebuildReglagesDynamiques();
            RebuildModulesIncertitude();
        }

        // ------- Masquage progressif des sections du formulaire -------
        // Le formulaire s'ouvre vide : seule la section Identification (N° FI) est visible.
        // Une fois un FI valide saisi, on dévoile Type/Nb/Module ; une fois le module choisi
        // (ou aucun module dispo pour ce type), on dévoile Source/Instrument/Mode.

        /// <summary>
        /// Étape 1 — N° FI au format <c>XX_NNNNN</c> (8 caractères, _ en position 3).
        /// Tant que false, toutes les sections suivantes restent masquées.
        /// </summary>
        public bool EtapeFIValide
        {
            get
            {
                var fi = MesureConfig.NumFI?.Trim() ?? string.Empty;
                return fi.Length == 8 && fi[2] == '_';
            }
        }

        /// <summary>
        /// Étape 2 — FI valide + module d'incertitude sélectionné (ou aucun dispo pour
        /// ce type de mesure, auquel cas le message d'aide remplace la sélection).
        /// Tant que false, les sections Source/Instrument/Mode restent masquées.
        /// </summary>
        public bool EtapeTypeValide =>
            EtapeFIValide && (ModuleSelectionne != null || !AModulesIncertitude);

        /// <summary>
        /// Visibilité de la section « Source du signal » : combinée avec
        /// <see cref="ShowSourceMesure"/> (Fréquence uniquement) ET l'étape 2.
        /// </summary>
        public bool AfficherSourceMesure => ShowSourceMesure && EtapeTypeValide;

        /// <summary>
        /// Notifie le XAML que les conditions de visibilité des sections ont changé.
        /// À appeler après toute modification de NumFI, TypeMesure, NbMesures ou Module.
        /// </summary>
        public void NotifierEtapes()
        {
            OnPropertyChanged(nameof(EtapeFIValide));
            OnPropertyChanged(nameof(EtapeTypeValide));
            OnPropertyChanged(nameof(AfficherSourceMesure));

            // Démarre le journal utilisateur de la FI dès que le numéro est complet/valide.
            // JournalFIService.DemarrerSession est idempotent : appels répétés avec la même
            // FI = no-op, donc on peut l'invoquer sans risque à chaque notification.
            if (EtapeFIValide)
            {
                string utilisateur = EtatApplication.UtilisateurConnecte?.NomComplet
                    ?? EtatApplication.UtilisateurConnecte?.Login
                    ?? "(inconnu)";
                string poste = EstSurBaie ? "Baie" : "Paillasse";
                JournalFIService.DemarrerSession(
                    MesureConfig.NumFI?.Trim() ?? string.Empty,
                    utilisateur,
                    poste);
            }
        }

        // ------- Modules d'incertitude (filtrés par TypeMesure) -------

        /// <summary>
        /// Liste des modules d'incertitude affichables dans la ComboBox "Module" — filtrés
        /// pour ne garder que ceux qui couvrent la <see cref="TypeMesure"/> courante (au
        /// moins une ligne CSV avec <c>Fonction</c> = équivalent du type).
        /// </summary>
        public ObservableCollection<ModuleIncertitude> ModulesDisponibles { get; } = new();

        private ModuleIncertitude? _moduleSelectionne;

        /// <summary>
        /// Module choisi par l'opérateur — son <c>NumModule</c> est sérialisé dans
        /// <see cref="MesureConfig"/> pour que l'<c>ExcelService</c> retrouve les coefficients
        /// CoeffA/CoeffB à la fin de la mesure.
        /// </summary>
        public ModuleIncertitude? ModuleSelectionne
        {
            get
            {
                if (_moduleSelectionne != null && ModulesDisponibles.Contains(_moduleSelectionne))
                    return _moduleSelectionne;
                return null;
            }
            set
            {
                if (ReferenceEquals(_moduleSelectionne, value)) return;
                _moduleSelectionne = value;
                MesureConfig.NumModuleIncertitude = value?.NumModule ?? string.Empty;
                OnPropertyChanged();
                NotifierEtapes();   // module sélectionné = passage à l'étape 3
            }
        }

        /// <summary>Vrai si au moins un module couvre le type de mesure courant.</summary>
        public bool AModulesIncertitude => ModulesDisponibles.Count > 0;

        // ------------------------------------------------------------------------------------
        // Module Fréquence auxiliaire (utilisé uniquement en Tachymètre Contact/Optique pour
        // alimenter ZNCoeffA / ZNCoeffB côté Hz, en complément du module tachy qui alimente
        // ZNCoeffC / ZNCoeffD côté RPM).
        // ------------------------------------------------------------------------------------
        public ObservableCollection<ModuleIncertitude> ModulesFreqDisponibles { get; } = new();

        private ModuleIncertitude? _moduleFreqSelectionne;
        public ModuleIncertitude? ModuleFreqSelectionne
        {
            get
            {
                if (_moduleFreqSelectionne != null && ModulesFreqDisponibles.Contains(_moduleFreqSelectionne))
                    return _moduleFreqSelectionne;
                return null;
            }
            set
            {
                if (ReferenceEquals(_moduleFreqSelectionne, value)) return;
                _moduleFreqSelectionne = value;
                MesureConfig.NumModuleIncertitudeFreq = value?.NumModule ?? string.Empty;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Vrai pour les types Tachymètre uniquement — déclenche l'affichage du second
        /// ComboBox « Module Fréquence (Hz) » dans la fenêtre Configuration.
        /// </summary>
        public bool ModuleFreqRequis =>
            EnTetesMesureHelper.EstTachymetre(MesureConfig.TypeMesure);

        /// <summary>
        /// Reconstruit la liste des modules selon <see cref="MesureConfig.TypeMesure"/>.
        /// Lit le sous-dossier dédié à ce type (ex. <c>Incertitudes\Frequence\</c>).
        /// Restaure la sélection persistée (<c>NumModuleIncertitude</c>) si elle est encore valide.
        /// Reconstruit également la liste des modules Fréquence auxiliaires (utilisée
        /// uniquement en tachymétrie).
        /// </summary>
        private void RebuildModulesIncertitude()
        {
            ModulesDisponibles.Clear();

            foreach (var m in ModulesIncertitudeService.Lister(MesureConfig.TypeMesure))
            {
                ModulesDisponibles.Add(m);
            }

            // Restaure la sélection précédente si toujours pertinente.
            string numPersist = MesureConfig?.NumModuleIncertitude ?? string.Empty;
            _moduleSelectionne = ModulesDisponibles.FirstOrDefault(m =>
                string.Equals(m.NumModule, numPersist, StringComparison.OrdinalIgnoreCase));
            if (_moduleSelectionne == null && MesureConfig != null)
            {
                MesureConfig.NumModuleIncertitude = string.Empty;
            }

            // Liste auxiliaire des modules Fréquence (uniquement pertinente pour tachy).
            ModulesFreqDisponibles.Clear();
            if (ModuleFreqRequis)
            {
                foreach (var m in ModulesIncertitudeService.Lister(TypeMesure.Frequence))
                    ModulesFreqDisponibles.Add(m);
            }
            string numFreqPersist = MesureConfig?.NumModuleIncertitudeFreq ?? string.Empty;
            _moduleFreqSelectionne = ModulesFreqDisponibles.FirstOrDefault(m =>
                string.Equals(m.NumModule, numFreqPersist, StringComparison.OrdinalIgnoreCase));
            if (_moduleFreqSelectionne == null && MesureConfig != null)
            {
                MesureConfig.NumModuleIncertitudeFreq = string.Empty;
            }

            OnPropertyChanged(nameof(ModuleSelectionne));
            OnPropertyChanged(nameof(AModulesIncertitude));
            OnPropertyChanged(nameof(ModuleFreqSelectionne));
            OnPropertyChanged(nameof(ModuleFreqRequis));
        }

        /// <summary>
        /// Déclenché quand <see cref="MesureConfig"/> est réassigné (typiquement par
        /// AcceuilViewModel qui passe la config persistée entre deux ouvertures de la fenêtre).
        /// Rebuild la liste des appareils et les réglages dynamiques pour restaurer l'état
        /// précédent (appareil sélectionné, options SCPI choisies) sans que l'utilisateur
        /// ait à tout re-saisir.
        /// </summary>
        partial void OnMesureConfigChanged(Mesure value)
        {
            RebuildAppareils();
            RebuildReglagesDynamiques();
            RebuildModulesIncertitude();
            RefreshAll();
        }

        /// <summary>
        /// Liste des appareils **détectés** sur le bus GPIB et disponibles pour être sélectionnés.
        /// Chaque appareil est marqué "catalogue" (prêt à l'emploi) ou "hors catalogue" (à enregistrer
        /// via Administration → Gérer les appareils avant utilisation).
        /// </summary>
        public ObservableCollection<OptionAppareil> Appareils { get; } = new();

        private void RebuildAppareils()
        {
            // Mémoriser la sélection courante pour la retrouver après reconstruction de la liste.
            var selFab = _appareilSelectionne?.Detecte?.Fabricant;
            var selMod = _appareilSelectionne?.Detecte?.Modele;

            // Re-synchronise ModeleReconnu pour chaque appareil détecté (le catalogue peut avoir changé).
            foreach (var det in EtatApplication.AppareilsDetectes)
            {
                det.ModeleReconnu = CatalogueAppareilsService.Instance.TrouverParIdn(det.Fabricant, det.Modele);
            }

            Appareils.Clear();

            // On n'affiche QUE les appareils réellement détectés sur le bus, pour éviter de
            // laisser l'utilisateur configurer une mesure sur un instrument qui n'est pas branché.
            foreach (var det in EtatApplication.AppareilsDetectes)
            {
                string suffixe = det.ModeleReconnu != null
                    ? $" ✓  (catalogue)"
                    : $" ✓  (hors catalogue — à enregistrer)";

                Appareils.Add(new OptionAppareil
                {
                    Libelle = $"{det.Libelle}{suffixe}",
                    Detecte = det
                });
            }

            // Retrouver la sélection courante — par IDN fabricant+modèle d'abord (cas d'un
            // simple refresh après scan), sinon par l'Id du modèle catalogue persisté dans
            // la Mesure (cas d'une réouverture de la fenêtre avec une config déjà validée).
            if (selFab != null || selMod != null)
            {
                _appareilSelectionne = Appareils.FirstOrDefault(o =>
                    o.Detecte != null
                    && o.Detecte.Fabricant == selFab
                    && o.Detecte.Modele == selMod);
            }
            else if (!string.IsNullOrEmpty(MesureConfig?.IdModeleCatalogue))
            {
                _appareilSelectionne = Appareils.FirstOrDefault(o =>
                    o.Detecte?.ModeleReconnu?.Id == MesureConfig.IdModeleCatalogue);
            }

            // Si rien n'a été restauré ET qu'il y a au moins un appareil détecté, on
            // sélectionne le premier — via la propriété pour que le setter renseigne
            // MesureConfig.IdModeleCatalogue. Sans ça, l'UI affichait bien le 1er appareil
            // (via le getter qui fait FirstOrDefault) mais l'orchestrator ne le voyait pas
            // côté config → erreur « Aucun appareil sélectionné » au lancement de la mesure.
            if (_appareilSelectionne == null && Appareils.Count > 0)
            {
                AppareilSelectionne = Appareils[0];
            }
            else if (_appareilSelectionne != null && MesureConfig != null)
            {
                // Restauré par IDN ou par IdModeleCatalogue : on resync la config au cas où
                // un changement de catalogue aurait modifié l'Id du modèle reconnu.
                var modele = _appareilSelectionne.Detecte?.ModeleReconnu;
                MesureConfig.IdModeleCatalogue = modele?.Id ?? string.Empty;
            }

            OnPropertyChanged(nameof(AppareilSelectionne));
            RebuildReglagesDynamiques();
            RefreshAll();
        }

        private OptionAppareil? _appareilSelectionne;

        /// <summary>
        /// Sélection courante dans la ComboBox. Pointe vers un appareil détecté sur le bus ; le
        /// modèle catalogue associé (si présent) est utilisé par l'orchestrator pour piloter la mesure.
        /// </summary>
        public OptionAppareil? AppareilSelectionne
        {
            get
            {
                if (_appareilSelectionne != null && Appareils.Contains(_appareilSelectionne))
                    return _appareilSelectionne;
                return Appareils.FirstOrDefault();
            }
            set
            {
                if (ReferenceEquals(_appareilSelectionne, value)) return;
                _appareilSelectionne = value;

                // Si le modèle catalogue est présent, on mémorise son Id pour l'orchestrator ;
                // sinon chaîne vide — l'orchestrator refusera de lancer et invitera l'utilisateur
                // à enregistrer l'appareil.
                var modele = ModeleCatalogueSelectionne();
                MesureConfig.IdModeleCatalogue = modele?.Id ?? string.Empty;

                OnPropertyChanged();
                OnPropertyChanged(nameof(MesureConfig));
                RebuildReglagesDynamiques();
                RefreshAll();
            }
        }

        /// <summary>
        /// Reconstruit la liste des réglages dynamiques + les sous-collections par voie/mode
        /// à partir du modèle catalogue correspondant à l'appareil sélectionné.
        /// </summary>
        private void RebuildReglagesDynamiques()
        {
            ReglagesDynamiques.Clear();
            ReglagesVoieA.Clear();
            ReglagesVoieB.Clear();
            ReglagesVoieC.Clear();
            ReglagesMode.Clear();

            var modele = ModeleCatalogueSelectionne();
            if (modele != null)
            {
                // Commandes SCPI persistées depuis la dernière validation de la Config.
                // Permet de restaurer la sélection précédente de l'utilisateur (impédance,
                // couplage, filtre, trigger, mode) quand il rouvre la fenêtre.
                var commandesPersistees = MesureConfig?.CommandesScpiReglages ?? new List<string>();

                foreach (var reglage in modele.Reglages)
                {
                    // Réglages "Auto" : ne s'affichent JAMAIS dans la fenêtre Configuration.
                    // L'option est sélectionnée automatiquement à la validation selon le contexte
                    // (TypeMesure + VoieActive), cf. CalculerCommandesAutomatiques.
                    if (reglage.Auto) continue;

                    var vm = new ReglageDynamiqueViewModel(reglage);

                    // Si une des commandes persistées correspond à une des options de ce réglage,
                    // on pré-sélectionne cette option au lieu de rester sur le défaut (1ère option).
                    var optionPersistee = reglage.Options
                        .FirstOrDefault(o => commandesPersistees.Any(c => c == o.CommandeScpi));
                    if (optionPersistee != null)
                    {
                        vm.OptionSelectionnee = optionPersistee;
                    }

                    ReglagesDynamiques.Add(vm);

                    // Dispatch dans la bonne sous-collection selon le nom canonique du réglage.
                    // Matching insensible aux accents est évité ici car les noms viennent de constantes.
                    if (reglage.Nom.Contains("Voie A")) ReglagesVoieA.Add(vm);
                    else if (reglage.Nom.Contains("Voie B")) ReglagesVoieB.Add(vm);
                    else if (reglage.Nom.Contains("Voie C")) ReglagesVoieC.Add(vm);
                    else ReglagesMode.Add(vm);
                }
            }

            OnPropertyChanged(nameof(AReglagesDynamiques));
            OnPropertyChanged(nameof(AReglagesVoieA));
            OnPropertyChanged(nameof(AReglagesVoieB));
            OnPropertyChanged(nameof(AReglagesVoieC));
            OnPropertyChanged(nameof(AReglagesMode));
            // La liste des temps de porte dépend du catalogue de l'appareil sélectionné.
            OnPropertyChanged(nameof(GateTimes));
            OnPropertyChanged(nameof(GateLibelleSelectionne));
            NotifierVoiesActives();
        }

        /// <summary>
        /// Retrouve le modèle du catalogue correspondant à l'appareil actuellement sélectionné
        /// (via l'IDN de la détection). Retourne null si pas détecté ou pas au catalogue.
        /// </summary>
        private ModeleAppareil? ModeleCatalogueSelectionne()
        {
            var det = AppareilSelectionne?.Detecte;
            if (det == null) return null;

            // 1) Priorité au matching direct déjà effectué lors du scan
            if (det.ModeleReconnu != null) return det.ModeleReconnu;

            // 2) Sinon on retente par IDN
            return CatalogueAppareilsService.Instance.TrouverParIdn(det.Fabricant, det.Modele);
        }

        public IEnumerable<TypeMesure> TypesMesure => Enum.GetValues(typeof(TypeMesure)).Cast<TypeMesure>();

        /// <summary>Échelle canonique complète des temps de porte (doit rester alignée avec
        /// <c>CatalogueAdapter._secondesSlotsUi</c> et <c>EnTetesMesureHelper._libellesGate</c>).</summary>
        private static readonly string[] _libellesGateCanoniques =
        {
            "10 ms", "20 ms", "50 ms", "100 ms", "200 ms", "500 ms",
            "1 s", "2 s", "5 s", "10 s", "20 s", "50 s",
            "100 s", "200 s", "500 s", "1000 s"
        };

        /// <summary>
        /// Temps de porte disponibles dans la ComboBox : **filtrés** sur ceux que l'utilisateur
        /// a cochés dans le catalogue pour l'appareil actuellement sélectionné. Si aucun appareil
        /// catalogue n'est sélectionné, renvoie la liste canonique complète.
        /// </summary>
        public List<string> GateTimes
        {
            get
            {
                var modele = ModeleCatalogueSelectionne();
                if (modele == null || modele.Gates.Count == 0)
                    return _libellesGateCanoniques.ToList();
                return modele.Gates.ToList();
            }
        }

        /// <summary>
        /// Libellé de la gate sélectionnée — binding <c>SelectedItem</c> de la ComboBox. La
        /// conversion vers <c>MesureConfig.GateIndex</c> (slot canonique 0..15) se fait ici
        /// pour que <c>AppliquerGateAsync</c> retrouve bien la commande SCPI correspondante
        /// dans le dictionnaire du catalogue.
        /// </summary>
        public string? GateLibelleSelectionne
        {
            get
            {
                int idx = MesureConfig?.GateIndex ?? -1;
                return (idx >= 0 && idx < _libellesGateCanoniques.Length)
                    ? _libellesGateCanoniques[idx]
                    : null;
            }
            set
            {
                if (value == null || MesureConfig == null) return;
                int slot = Array.IndexOf(_libellesGateCanoniques, value);
                if (slot < 0) return;
                if (MesureConfig.GateIndex == slot) return;
                MesureConfig.GateIndex = slot;
                OnPropertyChanged();
            }
        }

        public List<int> MeasurementCounts => Enumerable.Range(1, 100).ToList();

        /// <summary>
        /// Coefficients multiplicateurs proposés en mode Indirect. Index 0..4 correspond à
        /// l'exposant de 10 (×10⁰ … ×10⁴), utilisé par la formule Excel
        /// <c>POWER(10, ZNCoeffMult)</c> pour convertir la mesure en fréquence réelle.
        /// </summary>
        public List<string> CoefsMultiplicateurs => new()
        {
            "×1 (10⁰)", "×10 (10¹)", "×100 (10²)", "×1000 (10³)", "×10000 (10⁴)"
        };

        // Le temps de porte unique se configure ici uniquement pour les mesures de Fréquence.
        // Pour la Stabilité, la sélection se fait dans SelectionGateWindow (multi-gates à balayer
        // via cases à cocher + presets) — afficher un combo « TEMPS DE PORTE » ici serait
        // redondant et trompeur.
        public bool ShowGateSettings =>
            MesureConfig.TypeMesure == TypeMesure.Frequence ||
            MesureConfig.TypeMesure == TypeMesure.FreqAvantInterv ||
            MesureConfig.TypeMesure == TypeMesure.FreqFinale;

        /// <summary>
        /// Init manuelle disponible uniquement pour TypeMesure=Fréquence (conforme Delphi,
        /// TfrmConfigStanford.chkInitManu.Enabled := FMesure.TypeMesure = etFrequence).
        /// </summary>
        public bool InitManuDisponible => MesureConfig.TypeMesure == TypeMesure.Frequence;

        // Source du signal : visible seulement pour le type "Fréquence"
        public bool ShowSourceMesure => MesureConfig.TypeMesure == TypeMesure.Frequence;

        public bool IsSourceFrequencemetre
        {
            get => MesureConfig.SourceMesure == SourceMesure.Frequencemetre;
            set { if (value) { MesureConfig.SourceMesure = SourceMesure.Frequencemetre; RefreshAll(); } }
        }

        public bool IsSourceGenerateur
        {
            get => MesureConfig.SourceMesure == SourceMesure.Generateur;
            set { if (value) { MesureConfig.SourceMesure = SourceMesure.Generateur; RefreshAll(); } }
        }

        // Indirect disponible : pas en paillasse, pas intervalle / tachy / stroboscope
        public bool IndirectDisponible =>
            EstSurBaie
            && MesureConfig.TypeMesure != TypeMesure.Interval
            && MesureConfig.TypeMesure != TypeMesure.TachyContact
            && MesureConfig.TypeMesure != TypeMesure.TachyOptique
            && MesureConfig.TypeMesure != TypeMesure.Stroboscope;

        public bool IsModeDirect
        {
            get => MesureConfig.ModeMesure == ModeMesure.Direct;
            set { if (value) { MesureConfig.ModeMesure = ModeMesure.Direct; RefreshAll(); } }
        }

        public bool IsModeIndirect
        {
            get => MesureConfig.ModeMesure == ModeMesure.Indirect;
            set
            {
                if (value && !IndirectDisponible) return;
                if (value) { MesureConfig.ModeMesure = ModeMesure.Indirect; RefreshAll(); }
            }
        }

        public void RefreshAll()
        {
            OnPropertyChanged(nameof(IsModeDirect));
            OnPropertyChanged(nameof(IsModeIndirect));
            OnPropertyChanged(nameof(ShowGateSettings));
            OnPropertyChanged(nameof(ShowSourceMesure));
            OnPropertyChanged(nameof(IsSourceFrequencemetre));
            OnPropertyChanged(nameof(IsSourceGenerateur));
            OnPropertyChanged(nameof(IndirectDisponible));
            OnPropertyChanged(nameof(InitManuDisponible));
        }

        public void OnTypeMesureChanged()
        {
            if (MesureConfig.TypeMesure == TypeMesure.Interval)
                MesureConfig.NbMesures = 1;
            else
                MesureConfig.NbMesures = 30;

            // Si on quitte le type Fréquence, on repasse sur Fréquencemètre par défaut
            if (MesureConfig.TypeMesure != TypeMesure.Frequence)
                MesureConfig.SourceMesure = SourceMesure.Frequencemetre;

            // Si le mode indirect n'est plus dispo, on bascule sur Direct
            if (MesureConfig.ModeMesure == ModeMesure.Indirect && !IndirectDisponible)
                MesureConfig.ModeMesure = ModeMesure.Direct;

            // Si on quitte la Stabilité, on reset GateIndices à un seul élément (= la 1ère
            // gate de la sélection multi-gates). Sans ce reset, une mesure Fréquence
            // lancée juste après une Stabilité boucle sur toutes les gates qu'avait
            // sélectionnées la Stab — comportement non voulu.
            if (MesureConfig.TypeMesure != TypeMesure.Stabilite
                && MesureConfig.GateIndices.Count > 1)
            {
                MesureConfig.GateIndex = MesureConfig.GateIndices[0]; // setter remet la liste à 1 élément
            }

            // Le filtrage des modules d'incertitude dépend de TypeMesure → relister.
            RebuildModulesIncertitude();

            OnPropertyChanged(nameof(MesureConfig));
            RefreshAll();
            NotifierEtapes();   // changement type peut redéclencher AModulesIncertitude
        }

        public Action<bool>? CloseAction { get; set; }

        /// <summary>
        /// Valide la configuration : envoie d'abord les commandes SCPI des réglages dynamiques
        /// à l'appareil sélectionné, puis ferme la fenêtre.
        /// Si l'appareil n'est pas détecté ou sans réglages, ferme directement.
        ///
        /// Refuse de fermer tant que le numéro de fiche d'intervention n'est pas saisi — évite
        /// à l'utilisateur de devoir tout reparamétrer quand il lance une mesure et se rend
        /// compte qu'il a oublié le FI.
        /// </summary>
        [RelayCommand]
        private async Task ValiderAsync()
        {
            if (string.IsNullOrWhiteSpace(MesureConfig.NumFI))
            {
                MessageBox.Show(
                    "Le numéro de fiche d'intervention est obligatoire.\n\n"
                    + "Renseigne-le dans la section « Identification » avant de valider.",
                    "Numéro FI manquant",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            // ===== Portage exact de la validation Delphi (F_Configuration.pas:104-127) =====

            // 1. Trim + longueur exacte = 8 caractères
            string slNoFI = MesureConfig.NumFI.Trim();
            if (slNoFI.Length != 8)
            {
                MessageBox.Show(
                    $"N° de FI incorrect : {slNoFI} !",
                    "Erreur",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }
            MesureConfig.NumFI = slNoFI;   // normalise au cas où l'utilisateur avait laissé des espaces

            // 2. Vérification d'existence dans la base ASERi (SVR-OR / SIA / tAffaire)
            bool? existe = await AseriService.FiExisteAsync(slNoFI);
            string sFI = slNoFI.Replace('_', '/');

            if (existe == null)
            {
                // Erreur de connexion ASERi (réseau down, serveur en maintenance).
                // On laisse le choix à l'utilisateur : continuer sans vérif ou annuler.
                var resu = MessageBox.Show(
                    $"Impossible de joindre la base ASERi pour vérifier la FI n° {sFI}.\n\n"
                  + "Veux-tu continuer quand même (sans vérification d'existence) ?\n\n"
                  + "• Oui = lancer la mesure malgré tout (la FI sera marquée comme non vérifiée dans le journal).\n"
                  + "• Non = annuler et réessayer plus tard.",
                    "ASERi inaccessible",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);
                if (resu != MessageBoxResult.Yes) return;
                JournalLog.Warn(CategorieLog.Configuration, "ASERI_BYPASS_USER",
                    $"Validation FI {sFI} acceptée par l'utilisateur malgré l'indisponibilité d'ASERi.");
            }
            else if (existe == false)
            {
                // Fi inconnue de la base ASERi → refus identique au Delphi
                MessageBox.Show(
                    $"La FI n° {sFI} n'existe pas dans ASERi !",
                    "Erreur",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            await EnvoyerCommandesScpiAsync();
            CloseAction?.Invoke(true);
        }

        [RelayCommand] private void Annuler() => CloseAction?.Invoke(false);

        /// <summary>
        /// Calcule les commandes SCPI à envoyer automatiquement à l'appareil selon le
        /// contexte (TypeMesure + VoieActive). Lit les <see cref="ReglageAppareil"/> du
        /// catalogue qui ont le flag <c>Auto: true</c> et sélectionne pour chacun l'option
        /// dont le libellé matche le contexte.
        ///
        /// Convention de libellés reconnus (l'admin doit les respecter à l'enregistrement) :
        /// <list type="bullet">
        ///   <item>Mode de mesure :
        ///         "TIAB…" → Interval ;
        ///         "FREQ Voie A/B/C" → fréquence sur la voie correspondante</item>
        ///   <item>Résolution :
        ///         "CONT…" → Stabilité (Allan deviation gap-free) ;
        ///         "AUTO" → tout le reste (fallback par défaut)</item>
        /// </list>
        ///
        /// Si un libellé attendu n'est pas trouvé pour le contexte courant, on fallback sur
        /// la 1ère option du réglage (= valeur par défaut catalogue) — comportement permissif.
        /// </summary>
        private List<string> CalculerCommandesAutomatiques(string idModele)
        {
            _ = idModele; // gardé pour signature future, mais on lit le catalogue directement
            var commandes = new List<string>();
            var modele = ModeleCatalogueSelectionne();
            if (modele == null) return commandes;

            foreach (var reglage in modele.Reglages.Where(r => r.Auto))
            {
                var option = SelectionnerOptionAuto(reglage);
                if (option != null && !string.IsNullOrWhiteSpace(option.CommandeScpi))
                {
                    commandes.Add(option.CommandeScpi);
                }
            }

            return commandes;
        }

        /// <summary>
        /// Choisit l'option d'un réglage <c>Auto</c> selon le contexte courant (TypeMesure +
        /// VoieActive), par matching des libellés d'options.
        ///
        /// Conventions de libellés reconnus côté code :
        /// <list type="bullet">
        ///   <item>Mode de mesure : <c>"Voie A"</c> / <c>"Voie B"</c> / <c>"Voie C"</c> selon
        ///         <see cref="VoieActive"/> ; <c>"TIAB"</c> en priorité si
        ///         <see cref="TypeMesure.Interval"/>.</item>
        ///   <item>Résolution : <c>"CONT"</c> pour <see cref="TypeMesure.Stabilite"/>,
        ///         <c>"AUTO"</c> sinon (fallback).</item>
        /// </list>
        /// Si l'option attendue n'est pas trouvée, on prend la 1ère option du catalogue.
        ///
        /// Les champs <c>QuandType</c> / <c>QuandVoie</c> de <see cref="OptionReglage"/> sont
        /// disponibles mais non utilisés ici (les libellés suffisent pour les 2 cas connus —
        /// "Mode de mesure" et "Résolution" — qui sont les seuls réglages Auto en pratique).
        /// </summary>
        private OptionReglage? SelectionnerOptionAuto(ReglageAppareil reglage)
        {
            if (reglage.Options.Count == 0) return null;

            // Mode de mesure : Interval prioritaire (TIAB), sinon par voie active
            if (reglage.Nom.Contains("Mode", StringComparison.OrdinalIgnoreCase))
            {
                if (MesureConfig.TypeMesure == TypeMesure.Interval)
                {
                    var opt = reglage.Options.FirstOrDefault(o =>
                        o.Libelle.Contains("TIAB", StringComparison.OrdinalIgnoreCase));
                    if (opt != null) return opt;
                }
                string libelleVoie = MesureConfig.VoieActive switch
                {
                    VoieActive.A => "Voie A",
                    VoieActive.B => "Voie B",
                    VoieActive.C => "Voie C",
                    _ => "Voie A"
                };
                var optVoie = reglage.Options.FirstOrDefault(o =>
                    o.Libelle.Contains(libelleVoie, StringComparison.OrdinalIgnoreCase));
                if (optVoie != null) return optVoie;
            }

            // Résolution : CONT pour Stabilité, AUTO sinon
            if (reglage.Nom.Contains("Résolution", StringComparison.OrdinalIgnoreCase)
                || reglage.Nom.Contains("Resolution", StringComparison.OrdinalIgnoreCase))
            {
                if (MesureConfig.TypeMesure == TypeMesure.Stabilite)
                {
                    var opt = reglage.Options.FirstOrDefault(o =>
                        o.Libelle.Contains("CONT", StringComparison.OrdinalIgnoreCase));
                    if (opt != null) return opt;
                }
                var optAuto = reglage.Options.FirstOrDefault(o =>
                    o.Libelle.Contains("AUTO", StringComparison.OrdinalIgnoreCase));
                if (optAuto != null) return optAuto;
            }

            // Fallback : 1ère option (= défaut catalogue)
            return reglage.Options[0];
        }

        private async Task EnvoyerCommandesScpiAsync()
        {
            var modele = ModeleCatalogueSelectionne();
            var det = AppareilSelectionne?.Detecte;
            if (modele == null || det == null || ReglagesDynamiques.Count == 0)
            {
                MesureConfig.CommandesScpiReglages = new List<string>();
                return;
            }

            // Ne retient que les réglages de la voie active (+ le mode qui est global).
            // Ça évite de modifier les paramètres d'une voie que l'utilisateur n'utilise pas.
            var reglagesApplicables = new List<ReglageDynamiqueViewModel>();
            reglagesApplicables.AddRange(MesureConfig.VoieActive switch
            {
                VoieActive.A => ReglagesVoieA,
                VoieActive.B => ReglagesVoieB,
                VoieActive.C => ReglagesVoieC,
                _ => ReglagesVoieA
            });
            reglagesApplicables.AddRange(ReglagesMode);

            // Une CommandeScpi catalogue peut chaîner plusieurs commandes SCPI séparées par
            // ";:" (ex: ":INP1:LEV:AUTO OFF;:INP1:LEV {0}" sur le 53230A). On les split ici
            // en entrées distinctes pour que le driver émette UNE write GPIB par sous-commande.
            // Motivation : sur le firmware Keysight 53230A, certaines chaînes concaténées sont
            // tronquées au parsing — la seconde commande est silencieusement ignorée. Envoyer
            // séparément garantit que chaque ordre est bien pris en compte.
            var commandesUtilisateur = reglagesApplicables
                .Select(r => r.CommandeSelectionnee)
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .SelectMany(c => SplitCommandesScpi(c!))
                .ToList();

            // Commandes auto (typiquement Résolution :SENS:FREQ:MODE) selon TypeMesure + VoieActive.
            var commandesAuto = CalculerCommandesAutomatiques(modele.Id);

            // ⚠ ORDRE CRITIQUE pour le 53230A (et compteurs similaires).
            // Sur ces appareils, la commande "Mode de mesure" (CONF:FREQ / CONF:TINT / :FUNC)
            // réinitialise TOUS les paramètres d'entrée à leur valeur par défaut — y compris
            // INP:LEV:AUTO qui revient à ON et INP:LEV qui revient à 0V auto-calibré. Toute
            // commande INP:* envoyée AVANT CONF:* est donc silencieusement écrasée.
            //
            // Le catalogue marque "Mode de mesure" comme Auto=false (le user choisit Voie A/B/TIAB),
            // donc cette commande arrive dans le bucket `commandesUtilisateur`. On la sort ici
            // pour la placer en TÊTE de la séquence finale, suivie des autres commandes auto
            // (Résolution), puis seulement après les paramètres d'entrée et le trigger level.
            //
            // Ordre final :
            //   1. CONF:FREQ / CONF:TINT / :FUNC      (setup mesure — réinitialise tout)
            //   2. :SENS:FREQ:MODE AUTO|RECIPROCAL    (résolution, dépend du type de mesure)
            //   3. :INP1:IMP / :COUP / :FILT / :RANG  (paramètres d'entrée)
            //   4. :INP1:LEV:AUTO OFF puis :INP1:LEV  (trigger level — en DERNIER, pour ne
            //                                          surtout pas être écrasé par CONF:FREQ)
            bool EstCommandeSetup(string cmd) =>
                cmd.StartsWith("CONF:", StringComparison.OrdinalIgnoreCase) ||
                cmd.StartsWith(":FUNC", StringComparison.OrdinalIgnoreCase) ||
                cmd.StartsWith(":CONF", StringComparison.OrdinalIgnoreCase);

            var commandesSetup = commandesUtilisateur.Where(EstCommandeSetup).ToList();
            var commandesEntree = commandesUtilisateur.Where(c => !EstCommandeSetup(c)).ToList();

            var commandes = new List<string>();
            commandes.AddRange(commandesSetup);    // 1. CONF:FREQ (réinitialise les inputs)
            commandes.AddRange(commandesAuto);     // 2. :SENS:FREQ:MODE
            commandes.AddRange(commandesEntree);   // 3. INP1:IMP/COUP/FILT/RANG + LEV:AUTO OFF + LEV

            // On mémorise les commandes choisies dans la Mesure — l'orchestrator les rejouera
            // après le *RST pour que l'état de l'appareil soit cohérent avec la configuration
            // pendant la boucle de mesures.
            MesureConfig.CommandesScpiReglages = new List<string>(commandes);

            if (commandes.Count == 0) return;

            JournalLog.Info(CategorieLog.Configuration, "CONFIG_APPAREIL_DEBUT",
                $"Envoi de {commandes.Count} commande(s) à {modele.Nom} @ {det.AdresseCourte}.",
                new { modele.Nom, det.Adresse, Nb = commandes.Count });

            try
            {
                using var driver = new VisaIeeeDriver(det.Board);
                await AppareilScpiService.EnvoyerAsync(modele, det.Adresse, commandes, driver);

                JournalLog.Info(CategorieLog.Configuration, "CONFIG_APPAREIL_OK",
                    $"Configuration SCPI envoyée avec succès à {modele.Nom}.");
            }
            catch (Exception ex)
            {
                JournalLog.Erreur(CategorieLog.Configuration, "CONFIG_APPAREIL_ERR",
                    $"Échec de l'envoi SCPI à {modele.Nom} : {ex.Message}",
                    new { ex.GetType().Name, Commandes = commandes });

                MessageBox.Show(
                    $"Erreur lors de l'envoi des commandes à {modele.Nom} :\n\n{ex.Message}\n\n"
                    + "Les réglages n'ont pas été appliqués sur l'appareil — la configuration locale "
                    + "est quand même conservée.",
                    "Erreur de communication GPIB",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// Découpe une chaîne SCPI chaînée (séparateur ";:") en commandes individuelles bien
        /// formées (chacune préfixée par ":"). Exemple :
        /// <code>":INP1:LEV:AUTO OFF;:INP1:LEV 0" → [":INP1:LEV:AUTO OFF", ":INP1:LEV 0"]</code>
        /// Si la chaîne ne contient pas ";:", elle est renvoyée telle quelle (single-element).
        /// Utilisé pour garantir l'exécution séquentielle côté instrument sur les firmwares
        /// qui parsent mal la concaténation (53230A notamment).
        /// </summary>
        private static IEnumerable<string> SplitCommandesScpi(string commande)
        {
            if (string.IsNullOrWhiteSpace(commande)) yield break;
            if (!commande.Contains(";:"))
            {
                yield return commande;
                yield break;
            }
            var parts = commande.Split(";:", StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++)
            {
                string p = parts[i].Trim();
                if (string.IsNullOrEmpty(p)) continue;
                // La 1ʳᵉ partie garde son ":" initial déjà présent dans la chaîne d'origine.
                // Les suivantes ont perdu le ":" qui suivait le ";", on le réajoute pour
                // que la commande reparte de la racine SCPI.
                yield return i == 0 ? p : ":" + p;
            }
        }
    }
}
