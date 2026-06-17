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

        partial void OnEstSurBaieChanged(bool value)
        {
            OnPropertyChanged(nameof(ModeAdressesFixes));
            RebuildAppareils();
        }

        /// <summary>Vrai en poste Baie avec adresses fixes : affiche le champ d'adresse GPIB éditable.</summary>
        public bool ModeAdressesFixes => EstSurBaie && Metrologo.Models.EtatApplication.ModeAdressesFixes;

        /// <summary>Adresse GPIB saisie pour l'appareil legacy sélectionné (mode adresses fixes).
        /// Met aussi à jour MesureConfig.AdresseFixeForcee pour l'orchestrator.</summary>
        public int AdresseFixeSaisie
        {
            get => AppareilSelectionne?.AdresseFixe ?? 0;
            set
            {
                if (AppareilSelectionne == null) return;
                AppareilSelectionne.AdresseFixe = value;
                MesureConfig.AdresseFixeForcee = AppareilSelectionne.EstFixe ? value : -1;
                OnPropertyChanged();
            }
        }

        /// <summary>Réglages dynamiques du modèle sélectionné (impédance, couplage, filtre...),
        /// repeuplés depuis le catalogue à chaque changement d'appareil. Une ComboBox par réglage côté UI.</summary>
        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesDynamiques { get; } = new();

        // collections filtrées par voie, exposées au XAML

        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesVoieA { get; } = new();
        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesVoieB { get; } = new();
        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesVoieC { get; } = new();
        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesMode  { get; } = new();

        public bool AReglagesDynamiques => ReglagesDynamiques.Count > 0;
        public bool AReglagesVoieA => ReglagesVoieA.Count > 0;
        public bool AReglagesVoieB => ReglagesVoieB.Count > 0;
        public bool AReglagesVoieC => ReglagesVoieC.Count > 0;
        public bool AReglagesMode  => ReglagesMode.Count > 0;

        // sélection de la voie active (pilote la visibilité + le filtrage des commandes envoyées)

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

        /// <summary>Bloc Voie A visible dans le XAML : réglages définis ET voie active.</summary>
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

        // Masquage progressif du formulaire : à l'ouverture seule la section Identification
        // (N° FI) est visible. FI valide saisi -> on dévoile Type/Nb/Module ; module choisi
        // (ou aucun dispo pour ce type) -> on dévoile Source/Instrument/Mode.

        /// <summary>Étape 1 : N° FI au format XX_NNNNN (8 caractères, _ en position 3).
        /// Tant que false, tout le reste du formulaire est masqué.</summary>
        public bool EtapeFIValide
        {
            get
            {
                var fi = MesureConfig.NumFI?.Trim() ?? string.Empty;
                return fi.Length == 8 && fi[2] == '_';
            }
        }

        /// <summary>Étape 2 : FI valide + module d'incertitude choisi (ou aucun dispo pour ce type,
        /// auquel cas un message d'aide remplace la sélection). Tant que false, Source/Instrument/Mode restent masqués.</summary>
        public bool EtapeTypeValide =>
            EtapeFIValide && (ModuleSelectionne != null || !AModulesIncertitude);

        /// <summary>Section "Source du signal" : visible si type Fréquence (ShowSourceMesure) ET étape 2 validée.</summary>
        public bool AfficherSourceMesure => ShowSourceMesure && EtapeTypeValide;

        /// <summary>À appeler après modification de NumFI, TypeMesure, NbMesures ou Module
        /// pour rafraîchir la visibilité des sections.</summary>
        public void NotifierEtapes()
        {
            OnPropertyChanged(nameof(EtapeFIValide));
            OnPropertyChanged(nameof(EtapeTypeValide));
            OnPropertyChanged(nameof(AfficherSourceMesure));

            // démarre le journal utilisateur dès que le numéro de FI est valide.
            // DemarrerSession est idempotent (même FI = no-op), donc appelable
            // à chaque notification sans risque.
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

        /// <summary>Modules d'incertitude proposés dans la ComboBox "Module" : seulement ceux qui
        /// couvrent le TypeMesure courant (au moins une ligne CSV dont Fonction correspond au type).</summary>
        public ObservableCollection<ModuleIncertitude> ModulesDisponibles { get; } = new();

        private ModuleIncertitude? _moduleSelectionne;

        /// <summary>Module choisi par l'opérateur. Son NumModule est stocké dans MesureConfig
        /// pour que l'ExcelService retrouve les coefficients CoeffA/CoeffB en fin de mesure.</summary>
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

        // Module Fréquence auxiliaire : uniquement en Tachymètre Contact/Optique, alimente
        // ZNCoeffA/ZNCoeffB côté Hz (le module tachy alimente ZNCoeffC/ZNCoeffD côté RPM).
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

        /// <summary>Vrai pour les types Tachymètre : affiche le second ComboBox "Module Fréquence (Hz)".</summary>
        public bool ModuleFreqRequis =>
            EnTetesMesureHelper.EstTachymetre(MesureConfig.TypeMesure);

        /// <summary>Reliste les modules selon le TypeMesure (sous-dossier dédié, ex. Incertitudes\Frequence)
        /// et restaure la sélection persistée si encore valide. Idem pour la liste Fréquence auxiliaire (tachy).</summary>
        private void RebuildModulesIncertitude()
        {
            ModulesDisponibles.Clear();

            foreach (var m in ModulesIncertitudeService.Lister(MesureConfig.TypeMesure))
            {
                ModulesDisponibles.Add(m);
            }

            // restaure la sélection précédente si toujours pertinente
            string numPersist = MesureConfig?.NumModuleIncertitude ?? string.Empty;
            _moduleSelectionne = ModulesDisponibles.FirstOrDefault(m =>
                string.Equals(m.NumModule, numPersist, StringComparison.OrdinalIgnoreCase));
            if (_moduleSelectionne == null && MesureConfig != null)
            {
                MesureConfig.NumModuleIncertitude = string.Empty;
            }

            // liste auxiliaire des modules Fréquence (tachy uniquement)
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

        /// <summary>MesureConfig réassigné (AcceuilViewModel repasse la config persistée à la réouverture) :
        /// on rebuild appareils + réglages pour restaurer l'état précédent sans re-saisie.</summary>
        partial void OnMesureConfigChanged(Mesure value)
        {
            // On repart d'une sélection vierge pour que l'appareil soit restauré uniquement depuis
            // la valeur persistée (IdModeleCatalogue), pas depuis un _appareilSelectionne périmé.
            // Sinon en mode baie adresses fixes, la liste legacy n'étant jamais vide, le
            // constructeur auto-sélectionnait le 1er modèle qui écrasait la valeur persistée
            // (d'où le bug "le mode baie oublie l'appareil après une mesure" ; en mode classique
            // la liste détectée est vide à la construction, donc pas de souci).
            _appareilSelectionne = null;
            RebuildAppareils();
            RebuildReglagesDynamiques();
            RebuildModulesIncertitude();

            // Si la config restaurée est déjà en intervalle de temps, l'init manuelle doit
            // être cochée comme lors d'un switch (le Delphi forçait bInitManu à l'ouverture de
            // la fenêtre de config, cf. OnTypeMesureChanged / F_Main.pas:987).
            if (MesureConfig.TypeMesure == TypeMesure.Interval && !MesureConfig.InitManu)
                MesureConfig.InitManu = true;

            RefreshAll();
        }

        /// <summary>Appareils détectés sur le bus GPIB, marqués catalogue ou hors catalogue
        /// (à enregistrer via Administration avant utilisation).</summary>
        public ObservableCollection<OptionAppareil> Appareils { get; } = new();

        private void RebuildAppareils()
        {
            // mode adresses fixes (baie) : on liste les appareils legacy du catalogue plutôt
            // que ceux détectés sur le bus (les legacy ne répondent pas au *IDN?)
            if (EstSurBaie && Metrologo.Models.EtatApplication.ModeAdressesFixes)
            {
                RebuildAppareilsFixes();
                return;
            }

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

            // retrouve la sélection : par IDN fabricant+modèle d'abord (refresh après scan),
            // sinon par l'Id catalogue persisté dans la Mesure (réouverture de la fenêtre)
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

            // rien restauré : on sélectionne le premier, via la propriété pour que le setter
            // renseigne IdModeleCatalogue. Sans ça l'UI affichait le 1er appareil (getter
            // FirstOrDefault) mais la config restait vide -> "Aucun appareil sélectionné"
            // au lancement de la mesure.
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

        /// <summary>Liste des appareils en mode adresses fixes : tous les modèles legacy du catalogue,
        /// adresse GPIB pré-remplie éditable. La collision EIP/Stanford (16) est sans risque, un seul actif à la fois.</summary>
        private void RebuildAppareilsFixes()
        {
            string? selId = _appareilSelectionne?.ModeleFixe?.Id ?? MesureConfig?.IdModeleCatalogue;

            Appareils.Clear();
            foreach (var m in CatalogueAppareilsService.Instance.Modeles.Where(x => x.Parametres.Legacy))
            {
                Appareils.Add(new OptionAppareil
                {
                    Libelle = $"{m.Nom}  (adresse {m.Parametres.AdresseFixeParDefaut})",
                    ModeleFixe = m,
                    AdresseFixe = m.Parametres.AdresseFixeParDefaut
                });
            }

            _appareilSelectionne = Appareils.FirstOrDefault(o => o.ModeleFixe?.Id == selId);
            if (_appareilSelectionne == null && Appareils.Count > 0)
            {
                AppareilSelectionne = Appareils[0];   // via setter -> renseigne Id + AdresseFixeForcee
            }
            else if (_appareilSelectionne != null && MesureConfig != null)
            {
                MesureConfig.IdModeleCatalogue = _appareilSelectionne.ModeleFixe?.Id ?? string.Empty;
                MesureConfig.AdresseFixeForcee = _appareilSelectionne.AdresseFixe;
            }

            OnPropertyChanged(nameof(AppareilSelectionne));
            OnPropertyChanged(nameof(AdresseFixeSaisie));
            RebuildReglagesDynamiques();
            RefreshAll();
        }

        private OptionAppareil? _appareilSelectionne;

        /// <summary>Sélection courante dans la ComboBox ; le modèle catalogue associé sert à l'orchestrator.</summary>
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

                // Id du modèle catalogue pour l'orchestrator ; chaîne vide si hors catalogue
                // (l'orchestrator refusera alors de lancer la mesure)
                var modele = ModeleCatalogueSelectionne();
                MesureConfig.IdModeleCatalogue = modele?.Id ?? string.Empty;

                // Mode adresses fixes : transmet l'adresse forcée à l'orchestrator (sinon -1 = IDN).
                MesureConfig.AdresseFixeForcee =
                    _appareilSelectionne?.EstFixe == true ? _appareilSelectionne.AdresseFixe : -1;

                OnPropertyChanged();
                OnPropertyChanged(nameof(MesureConfig));
                OnPropertyChanged(nameof(AdresseFixeSaisie));
                RebuildReglagesDynamiques();
                RefreshAll();
            }
        }

        /// <summary>Reconstruit les réglages dynamiques et les sous-collections par voie/mode
        /// depuis le modèle catalogue de l'appareil sélectionné.</summary>
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
                    // réglages Auto : jamais affichés ici, l'option est choisie à la validation
                    // selon TypeMesure + VoieActive (cf. CalculerCommandesAutomatiques).
                    // "Mode de mesure" est traité pareil : plus de menu déroulant, CONF:FREQ
                    // est toujours calculé d'après la voie active.
                    if (reglage.Auto || EstReglageMode(reglage)) continue;

                    var vm = new ReglageDynamiqueViewModel(reglage);

                    // Restaure la sélection précédente de l'utilisateur (Choix OU valeur numérique
                    // comme le Trigger) à partir des commandes SCPI persistées.
                    vm.RestaurerDepuis(commandesPersistees);

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
            OnPropertyChanged(nameof(ShowReglagesDynamiques));
            OnPropertyChanged(nameof(AReglagesVoieA));
            OnPropertyChanged(nameof(AReglagesVoieB));
            OnPropertyChanged(nameof(AReglagesVoieC));
            OnPropertyChanged(nameof(AReglagesMode));
            // La liste des temps de porte dépend du catalogue de l'appareil sélectionné.
            OnPropertyChanged(nameof(GateTimes));
            OnPropertyChanged(nameof(GateLibelleSelectionne));
            NotifierVoiesActives();
        }

        /// <summary>Vrai si c'est le réglage "Mode de mesure" (CONF:FREQ / :FUNC) : masqué de l'UI,
        /// sa commande est toujours calculée d'après la voie active (cf. SelectionnerOptionAuto).</summary>
        private static bool EstReglageMode(ReglageAppareil reglage)
            => reglage.Nom.Contains("Mode", StringComparison.OrdinalIgnoreCase);

        /// <summary>Modèle catalogue de l'appareil sélectionné (via l'IDN), null si pas détecté ou hors catalogue.</summary>
        private ModeleAppareil? ModeleCatalogueSelectionne()
        {
            // Mode adresses fixes : le modèle legacy est porté directement par l'option sélectionnée.
            if (AppareilSelectionne?.ModeleFixe != null) return AppareilSelectionne.ModeleFixe;

            var det = AppareilSelectionne?.Detecte;
            if (det == null) return null;

            // 1) Priorité au matching direct déjà effectué lors du scan
            if (det.ModeleReconnu != null) return det.ModeleReconnu;

            // 2) Sinon on retente par IDN
            return CatalogueAppareilsService.Instance.TrouverParIdn(det.Fabricant, det.Modele);
        }

        public IEnumerable<TypeMesure> TypesMesure => Enum.GetValues(typeof(TypeMesure)).Cast<TypeMesure>();

        /// <summary>Échelle canonique des temps de porte, à garder alignée avec
        /// CatalogueAdapter._secondesSlotsUi et EnTetesMesureHelper._libellesGate.</summary>
        private static readonly string[] _libellesGateCanoniques =
        {
            "10 ms", "20 ms", "50 ms", "100 ms", "200 ms", "500 ms",
            "1 s", "2 s", "5 s", "10 s", "20 s", "50 s",
            "100 s", "200 s", "500 s", "1000 s"
        };

        /// <summary>Temps de porte de la ComboBox, filtrés sur ceux cochés au catalogue pour
        /// l'appareil sélectionné. Liste canonique complète si pas d'appareil catalogue.</summary>
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

        /// <summary>Libellé de la gate sélectionnée (SelectedItem de la ComboBox). La conversion vers
        /// GateIndex (slot canonique 0..15) se fait ici pour qu'AppliquerGateAsync retrouve la commande SCPI au catalogue.</summary>
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

        /// <summary>Multiplicateurs du mode Indirect. L'index 0..4 = exposant de 10, utilisé par la
        /// formule Excel POWER(10, ZNCoeffMult) pour convertir la mesure en fréquence réelle.</summary>
        public List<string> CoefsMultiplicateurs => new()
        {
            "×1 (10⁰)", "×10 (10¹)", "×100 (10²)", "×1000 (10³)", "×10000 (10⁴)"
        };

        // le temps de porte unique ne se configure ici que pour la Fréquence ; en Stabilité
        // la sélection multi-gates se fait dans SelectionGateWindow, un combo ici serait redondant
        public bool ShowGateSettings =>
            MesureConfig.TypeMesure == TypeMesure.Frequence ||
            MesureConfig.TypeMesure == TypeMesure.FreqAvantInterv ||
            MesureConfig.TypeMesure == TypeMesure.FreqFinale;

        /// <summary>Init manuelle proposée en Fréquence et en Intervalle de temps. En intervalle
        /// elle est forcée cochée à l'entrée (cf. OnTypeMesureChanged, conforme Delphi
        /// F_Main.pas:987 / 1025-1027 pour Racal-Dana et Stanford), mais reste visible et
        /// décochable par l'opérateur.</summary>
        public bool InitManuDisponible =>
            MesureConfig.TypeMesure == TypeMesure.Frequence
            || MesureConfig.TypeMesure == TypeMesure.Interval;

        /// <summary>Case "Initialisation manuelle" (wrapper sur le modèle pour notifier l'UI).
        /// Cochée : l'opérateur configure l'appareil à la main, on n'envoie rien et on masque les réglages.</summary>
        public bool InitManu
        {
            get => MesureConfig.InitManu;
            set
            {
                if (MesureConfig.InitManu == value) return;
                MesureConfig.InitManu = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ShowReglagesDynamiques));
            }
        }

        /// <summary>Réglages dynamiques affichés seulement si l'appareil en propose et que l'init
        /// manuelle n'est pas cochée (en init manuelle l'opérateur paramètre l'appareil lui-même).</summary>
        public bool ShowReglagesDynamiques => AReglagesDynamiques && !MesureConfig.InitManu;

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
            OnPropertyChanged(nameof(InitManu));
            OnPropertyChanged(nameof(ShowReglagesDynamiques));
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
            // sélectionnées la Stab, comportement non voulu.
            if (MesureConfig.TypeMesure != TypeMesure.Stabilite
                && MesureConfig.GateIndices.Count > 1)
            {
                MesureConfig.GateIndex = MesureConfig.GateIndices[0]; // setter remet la liste à 1 élément
            }

            // En intervalle de temps, l'init manuelle est forcée cochée dès qu'on bascule
            // sur ce type (conforme Delphi F_Main.pas:987 / 1025-1027 :
            // bInitManu := (TypeMesure = etInterval), aussi bien pour le Racal-Dana que pour
            // le Stanford). La case reste visible et décochable si l'opérateur préfère piloter
            // l'init lui-même. Pour les autres types où l'init manuelle n'est pas disponible,
            // on la décoche pour qu'une case masquée ne reste pas active.
            if (MesureConfig.TypeMesure == TypeMesure.Interval)
                MesureConfig.InitManu = true;
            else if (!InitManuDisponible && MesureConfig.InitManu)
                MesureConfig.InitManu = false;

            // le filtrage des modules d'incertitude dépend de TypeMesure -> relister
            RebuildModulesIncertitude();

            OnPropertyChanged(nameof(MesureConfig));
            RefreshAll();
            NotifierEtapes();   // changement type peut redéclencher AModulesIncertitude
        }

        public Action<bool>? CloseAction { get; set; }

        /// <summary>Valide la config : envoie les commandes SCPI des réglages à l'appareil puis ferme.
        /// Refuse de fermer sans numéro de FI, pour éviter de devoir tout reparamétrer après coup.</summary>
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
                // FI inconnue de la base ASERi -> refus identique au Delphi
                MessageBox.Show(
                    $"La FI n° {sFI} n'existe pas dans ASERi !",
                    "Erreur",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            if (MesureConfig.InitManu)
            {
                // Init manuelle : l'opérateur a configuré l'appareil à la main (impédance,
                // couplage, mode…). On n'envoie aucune commande de configuration et on vide
                // les réglages pour que rien ne soit rejoué après un éventuel *RST (lequel est
                // de toute façon sauté côté orchestrateur en init manuelle).
                MesureConfig.CommandesScpiReglages = new List<string>();
            }
            else
            {
                await EnvoyerCommandesScpiAsync();
            }

            // En intervalle de temps, l'appareil est configuré à la main (init manuelle) :
            // Metrologo n'envoie aucune commande. Dernier rappel avant le lancement pour que
            // l'opérateur vérifie les réglages du fréquencemètre. OK lance la mesure.
            if (MesureConfig.TypeMesure == TypeMesure.Interval)
            {
                MessageBox.Show(
                    "Vérifie les paramètres du fréquencemètre avant de lancer la mesure.\n\n"
                    + "En intervalle de temps, l'appareil est configuré à la main : "
                    + "aucune commande n'est envoyée par Metrologo.",
                    "Vérification des paramètres",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }

            CloseAction?.Invoke(true);
        }

        [RelayCommand] private void Annuler() => CloseAction?.Invoke(false);

        /// <summary>Commandes SCPI envoyées automatiquement selon le contexte (TypeMesure + VoieActive) :
        /// pour chaque réglage Auto du catalogue, choisit l'option dont le libellé matche (cf. SelectionnerOptionAuto).</summary>
        private List<string> CalculerCommandesAutomatiques(string idModele)
        {
            _ = idModele; // gardé pour signature future, mais on lit le catalogue directement
            var commandes = new List<string>();
            var modele = ModeleCatalogueSelectionne();
            if (modele == null) return commandes;

            // "Mode de mesure" est inclus même s'il n'est pas marqué Auto au catalogue :
            // il n'est plus proposé à l'utilisateur, sa commande CONF:FREQ est toujours
            // calculée automatiquement selon la voie active.
            foreach (var reglage in modele.Reglages.Where(r => r.Auto || EstReglageMode(r)))
            {
                var option = SelectionnerOptionAuto(reglage);
                if (option != null && !string.IsNullOrWhiteSpace(option.CommandeScpi))
                {
                    commandes.Add(option.CommandeScpi);
                }
            }

            return commandes;
        }

        /// <summary>Choisit l'option d'un réglage Auto par matching de libellés : "TIAB" si Interval, "Voie A/B/C"
        /// selon la voie active, "CONT" en Stabilité, "AUTO" sinon ; à défaut la 1ère option du catalogue.
        /// Les champs QuandType/QuandVoie ne sont pas utilisés, les libellés suffisent en pratique.</summary>
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
            if (modele == null || ReglagesDynamiques.Count == 0)
            {
                MesureConfig.CommandesScpiReglages = new List<string>();
                return;
            }

            // Adresse cible : appareil détecté sur le bus (scan) OU adresse fixe (appareil legacy
            // sans *IDN?). Sans cette branche, le mode adresses fixes laissait
            // CommandesScpiReglages vide -> impédance/couplage jamais rejoués par l'orchestrator.
            var det = AppareilSelectionne?.Detecte;
            bool estFixe = AppareilSelectionne?.EstFixe == true;
            if (det == null && !estFixe)
            {
                MesureConfig.CommandesScpiReglages = new List<string>();
                return;
            }
            int board = det?.Board ?? 0;
            int adresse = det?.Adresse ?? AppareilSelectionne!.AdresseFixe;
            string adresseCourte = det?.AdresseCourte ?? $"GPIB{board}::{adresse}";

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

            // une CommandeScpi catalogue peut chaîner plusieurs commandes séparées par ";:"
            // (ex ":INP1:LEV:AUTO OFF;:INP1:LEV {0}" sur le 53230A). On split en writes GPIB
            // distinctes : le firmware du 53230A tronque parfois les chaînes concaténées et
            // ignore silencieusement la seconde commande.
            var commandesUtilisateur = reglagesApplicables
                .Select(r => r.CommandeSelectionnee)
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .SelectMany(c => SplitCommandesScpi(c!))
                .ToList();

            // Commandes auto (typiquement Résolution :SENS:FREQ:MODE) selon TypeMesure + VoieActive.
            var commandesAuto = CalculerCommandesAutomatiques(modele.Id);

            // ORDRE CRITIQUE pour le 53230A (et compteurs similaires) : CONF:FREQ / CONF:TINT /
            // :FUNC réinitialise TOUS les paramètres d'entrée (INP:LEV:AUTO repasse ON, INP:LEV
            // revient à 0V). Toute commande INP:* envoyée avant CONF:* est donc écrasée.
            // Ordre final :
            //   1. CONF:FREQ / CONF:TINT / :FUNC      (setup mesure, réinitialise tout)
            //   2. :SENS:FREQ:MODE AUTO|RECIPROCAL    (résolution, dépend du type de mesure)
            //   3. :INP1:IMP / :COUP / :FILT / :RANG  (paramètres d'entrée)
            //   4. :INP1:LEV:AUTO OFF puis :INP1:LEV  (trigger level en dernier, pour ne pas
            //                                          être écrasé par CONF:FREQ)
            bool EstCommandeSetup(string cmd) =>
                cmd.StartsWith("CONF:", StringComparison.OrdinalIgnoreCase) ||
                cmd.StartsWith(":FUNC", StringComparison.OrdinalIgnoreCase) ||
                cmd.StartsWith(":CONF", StringComparison.OrdinalIgnoreCase);

            // CONF:FREQ / :FUNC vient maintenant des commandes auto ("Mode de mesure" n'est plus
            // un réglage utilisateur) : on extrait le setup de toutes les sources pour le garder en tête.
            var commandesSetup = commandesUtilisateur.Where(EstCommandeSetup)
                .Concat(commandesAuto.Where(EstCommandeSetup)).ToList();
            var commandesAutoReste = commandesAuto.Where(c => !EstCommandeSetup(c)).ToList();
            var commandesEntree = commandesUtilisateur.Where(c => !EstCommandeSetup(c)).ToList();

            var commandes = new List<string>();
            commandes.AddRange(commandesSetup);       // 1. CONF:FREQ / :FUNC (réinitialise les inputs)
            commandes.AddRange(commandesAutoReste);   // 2. :SENS:FREQ:MODE (résolution)
            commandes.AddRange(commandesEntree);      // 3. INP1:IMP/COUP/FILT/RANG + LEV:AUTO OFF + LEV

            // commandes mémorisées dans la Mesure : l'orchestrator les rejouera après le *RST
            // pour garder l'appareil cohérent avec la config pendant la boucle de mesures
            MesureConfig.CommandesScpiReglages = new List<string>(commandes);

            if (commandes.Count == 0) return;

            JournalLog.Info(CategorieLog.Configuration, "CONFIG_APPAREIL_DEBUT",
                $"Envoi de {commandes.Count} commande(s) à {modele.Nom} @ {adresseCourte}.",
                new { modele.Nom, Adresse = adresse, Nb = commandes.Count });

            try
            {
                using var driver = new VisaIeeeDriver(board);
                await AppareilScpiService.EnvoyerAsync(modele, adresse, commandes, driver);

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

        /// <summary>Découpe une chaîne SCPI chaînée par ";:" en commandes individuelles re-préfixées
        /// par ":". Nécessaire pour les firmwares qui parsent mal la concaténation (53230A notamment).</summary>
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
                // la 1ère partie garde son ":" initial ; les suivantes ont perdu le ":" du
                // séparateur, on le remet pour repartir de la racine SCPI
                yield return i == 0 ? p : ":" + p;
            }
        }
    }
}
