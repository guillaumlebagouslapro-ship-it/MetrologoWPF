using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
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

        /// <summary>Adresse GPIB de l'appareil legacy sélectionné. Met aussi à jour MesureConfig.AdresseFixeForcee.</summary>
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

        /// <summary>Réglages dynamiques (impédance, couplage, filtre…), rechargés depuis le catalogue à chaque changement d'appareil.</summary>
        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesDynamiques { get; } = new();

        // Sous-collections filtrées par voie, pour les bindings XAML.

        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesVoieA { get; } = new();
        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesVoieB { get; } = new();
        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesVoieC { get; } = new();
        public ObservableCollection<ReglageDynamiqueViewModel> ReglagesMode  { get; } = new();

        public bool AReglagesDynamiques => ReglagesDynamiques.Count > 0;
        public bool AReglagesVoieA => ReglagesVoieA.Count > 0;
        public bool AReglagesVoieB => ReglagesVoieB.Count > 0;
        public bool AReglagesVoieC => ReglagesVoieC.Count > 0;
        public bool AReglagesMode  => ReglagesMode.Count > 0;

        // Voie active : pilote la visibilité et le filtrage des commandes SCPI envoyées.

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

        // Dévoilement progressif : FI valide -> Type/Nb/Module ; module choisi (ou aucun dispo) -> Source/Instrument/Mode.

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

        /// <summary>Étape 2 : FI valide + module choisi (ou aucun dispo). Tant que false, Source/Instrument/Mode restent masqués.</summary>
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

            // DemarrerSession est idempotent (même FI = no-op), donc appelable à chaque notification.
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

        /// <summary>Modules d'incertitude pour la ComboBox "Module", filtrés sur le TypeMesure courant.</summary>
        public ObservableCollection<ModuleIncertitude> ModulesDisponibles { get; } = new();

        private ModuleIncertitude? _moduleSelectionne;

        /// <summary>Module choisi par l'opérateur. NumModule stocké dans MesureConfig pour que l'ExcelService retrouve CoeffA/CoeffB.</summary>
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

        // Module Fréquence auxiliaire (tachy uniquement) : alimente ZNCoeffA/ZNCoeffB (Hz) ;
        // le module tachy alimente ZNCoeffC/ZNCoeffD (RPM).
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

        /// <summary>Reliste les modules pour le TypeMesure courant et restaure la sélection persistée si valide.
        /// Idem pour la liste Fréquence auxiliaire (tachy).</summary>
        private void RebuildModulesIncertitude()
        {
            ModulesDisponibles.Clear();

            foreach (var m in ModulesIncertitudeService.Lister(MesureConfig.TypeMesure))
            {
                ModulesDisponibles.Add(m);
            }

            // restaure la sélection précédente si elle est toujours dans la liste
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

        /// <summary>MesureConfig réassigné (réouverture) : rebuild pour restaurer l'état sans re-saisie.</summary>
        partial void OnMesureConfigChanged(Mesure value)
        {
            // Repart d'une sélection vierge pour restaurer l'appareil depuis IdModeleCatalogue,
            // pas depuis un _appareilSelectionne périmé. Sans ça, en mode baie, le 1er legacy
            // était auto-sélectionné et écrasait la valeur persistée (bug "oublie l'appareil après une mesure").
            _appareilSelectionne = null;
            RebuildAppareils();
            RebuildReglagesDynamiques();
            RebuildModulesIncertitude();

            // Si la config restaurée est en intervalle, force l'init manuelle (cf. F_Main.pas:987).
            if (MesureConfig.TypeMesure == TypeMesure.Interval && !MesureConfig.InitManu)
                MesureConfig.InitManu = true;

            SyncIntervalleTexte();   // aligne les champs texte sur la config restaurée
            RefreshAll();
        }

        /// <summary>Appareils détectés sur le bus GPIB (catalogue ou hors catalogue — à enregistrer avant utilisation).</summary>
        public ObservableCollection<OptionAppareil> Appareils { get; } = new();

        private void RebuildAppareils()
        {
            // Mode baie adresses fixes : liste les legacy du catalogue (ils ne répondent pas au *IDN?).
            if (EstSurBaie && Metrologo.Models.EtatApplication.ModeAdressesFixes)
            {
                RebuildAppareilsFixes();
                return;
            }

            // Mémorise la sélection pour la retrouver après reconstruction.
            var selFab = _appareilSelectionne?.Detecte?.Fabricant;
            var selMod = _appareilSelectionne?.Detecte?.Modele;

                // Resynchronise ModeleReconnu (le catalogue peut avoir changé depuis le dernier scan).
            foreach (var det in EtatApplication.AppareilsDetectes)
            {
                det.ModeleReconnu = CatalogueAppareilsService.Instance.TrouverParIdn(det.Fabricant, det.Modele);
            }

            Appareils.Clear();

            // Seuls les appareils détectés sur le bus — pas de config sur un instrument absent.
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

            // Retrouve la sélection : par IDN (refresh après scan) ou par IdModeleCatalogue (réouverture).
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

            // Rien restauré : sélectionne le 1er via la propriété (setter renseigne IdModeleCatalogue).
            // Sans ça, l'UI affichait le 1er appareil mais la config restait vide -> "Aucun appareil" au lancement.
            if (_appareilSelectionne == null && Appareils.Count > 0)
            {
                AppareilSelectionne = Appareils[0];
            }
            else if (_appareilSelectionne != null && MesureConfig != null)
            {
                // Resync l'Id au cas où un changement de catalogue aurait modifié le modèle reconnu.
                var modele = _appareilSelectionne.Detecte?.ModeleReconnu;
                MesureConfig.IdModeleCatalogue = modele?.Id ?? string.Empty;
            }

            OnPropertyChanged(nameof(AppareilSelectionne));
            RebuildReglagesDynamiques();
            RefreshAll();
        }

        /// <summary>Mode adresses fixes : liste tous les legacy du catalogue, adresse GPIB éditable.
        /// Collision EIP/Stanford (addr 16) sans risque : un seul actif à la fois.</summary>
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
                AppareilSelectionne = Appareils[0];   // setter -> renseigne Id + AdresseFixeForcee
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

        /// <summary>Appareil sélectionné dans la ComboBox ; le modèle catalogue associé est utilisé par l'orchestrator.</summary>
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

                // Id catalogue pour l'orchestrator ; vide si hors catalogue (mesure refusée).
                var modele = ModeleCatalogueSelectionne();
                MesureConfig.IdModeleCatalogue = modele?.Id ?? string.Empty;

                // Adresse fixe transmise à l'orchestrator (-1 si IDN).
                MesureConfig.AdresseFixeForcee =
                    _appareilSelectionne?.EstFixe == true ? _appareilSelectionne.AdresseFixe : -1;

                OnPropertyChanged();
                OnPropertyChanged(nameof(MesureConfig));
                OnPropertyChanged(nameof(AdresseFixeSaisie));
                RebuildReglagesDynamiques();
                RefreshAll();
            }
        }

        /// <summary>Reconstruit les réglages dynamiques et les sous-collections par voie/mode depuis le catalogue.</summary>
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
                // Commandes persistées depuis la dernière validation : restaure la sélection (impédance, couplage, filtre…).
                var commandesPersistees = MesureConfig?.CommandesScpiReglages ?? new List<string>();

                foreach (var reglage in modele.Reglages)
                {
                    // Réglages Auto et "Mode de mesure" masqués ici : option calculée à la validation
                    // selon TypeMesure + VoieActive (cf. CalculerCommandesAutomatiques).
                    if (reglage.Auto || EstReglageMode(reglage)) continue;

                    var vm = new ReglageDynamiqueViewModel(reglage);

                    // Restaure la sélection (Choix ou valeur numérique comme le Trigger) depuis les commandes persistées.
                    vm.RestaurerDepuis(commandesPersistees);

                    ReglagesDynamiques.Add(vm);

                    // Dispatch dans la sous-collection selon le nom canonique (noms de constantes, pas de risque d'accents).
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
            // GateTimes dépend du catalogue de l'appareil sélectionné.
            OnPropertyChanged(nameof(GateTimes));
            OnPropertyChanged(nameof(GateLibelleSelectionne));
            NotifierVoiesActives();
        }

        /// <summary>Vrai si c'est le réglage "Mode de mesure" (CONF:FREQ / :FUNC) : masqué de l'UI,
        /// sa commande est toujours calculée d'après la voie active (cf. SelectionnerOptionAuto).</summary>
        private static bool EstReglageMode(ReglageAppareil reglage)
            => reglage.Nom.Contains("Mode", StringComparison.OrdinalIgnoreCase);

        /// <summary>Modèle catalogue de l'appareil sélectionné, null si absent ou hors catalogue.</summary>
        private ModeleAppareil? ModeleCatalogueSelectionne()
        {
            // Mode adresses fixes : le modèle legacy est porté par l'option elle-même.
            if (AppareilSelectionne?.ModeleFixe != null) return AppareilSelectionne.ModeleFixe;

            var det = AppareilSelectionne?.Detecte;
            if (det == null) return null;

            // Matching direct du scan, sinon retente par IDN.
            if (det.ModeleReconnu != null) return det.ModeleReconnu;
            return CatalogueAppareilsService.Instance.TrouverParIdn(det.Fabricant, det.Modele);
        }

        public IEnumerable<TypeMesure> TypesMesure => Enum.GetValues(typeof(TypeMesure)).Cast<TypeMesure>();

        /// <summary>Echelle canonique des gates — à garder alignée avec CatalogueAdapter._secondesSlotsUi et EnTetesMesureHelper._libellesGate.</summary>
        private static readonly string[] _libellesGateCanoniques =
        {
            "10 ms", "20 ms", "50 ms", "100 ms", "200 ms", "500 ms",
            "1 s", "2 s", "5 s", "10 s", "20 s", "50 s",
            "100 s", "200 s", "500 s", "1000 s"
        };

        /// <summary>Gates disponibles : celles cochées au catalogue pour l'appareil, ou la liste complète si hors catalogue.</summary>
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

        /// <summary>Libellé de la gate sélectionnée. Converti en GateIndex (0..15) pour que l'orchestrator retrouve la commande SCPI au catalogue.</summary>
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

        /// <summary>Multiplicateurs du mode Indirect. Index 0..4 = exposant de 10 (POWER(10, ZNCoeffMult) dans Excel).</summary>
        public List<string> CoefsMultiplicateurs => new()
        {
            "×1 (10⁰)", "×10 (10¹)", "×100 (10²)", "×1000 (10³)", "×10000 (10⁴)"
        };

        // Gate unique : visible en Fréquence seulement. En Stabilité la sélection multi-gates est dans SelectionGateWindow.
        public bool ShowGateSettings =>
            MesureConfig.TypeMesure == TypeMesure.Frequence ||
            MesureConfig.TypeMesure == TypeMesure.FreqAvantInterv ||
            MesureConfig.TypeMesure == TypeMesure.FreqFinale;

        /// <summary>Init manuelle disponible en Fréquence et Intervalle. En intervalle, forcée cochée à l'entrée
        /// (conforme Delphi F_Main.pas:987), mais décochable par l'opérateur.</summary>
        public bool InitManuDisponible =>
            MesureConfig.TypeMesure == TypeMesure.Frequence
            || MesureConfig.TypeMesure == TypeMesure.Interval;

        /// <summary>Init manuelle cochée : l'opérateur configure l'appareil à la main, aucune commande envoyée, réglages masqués.</summary>
        public bool InitManu
        {
            get => MesureConfig.InitManu;
            set
            {
                if (MesureConfig.InitManu == value) return;
                MesureConfig.InitManu = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ShowReglagesDynamiques));
                NotifierIntervalle();   // décocher l'init manuelle révèle le panneau intervalle
            }
        }

        /// <summary>Réglages dynamiques affichés si l'appareil en propose, sans init manuelle, et hors intervalle
        /// (l'intervalle a son propre panneau, cf. ShowIntervalleConfig).</summary>
        public bool ShowReglagesDynamiques =>
            AReglagesDynamiques
            && !MesureConfig.InitManu
            && MesureConfig.TypeMesure != TypeMesure.Interval;

        // ===== Panneau Intervalle de temps (53230A) — visible si Intervalle + init manuelle décochée. =====

        /// <summary>Panneau intervalle affiché : type Intervalle + init manuelle décochée + appareil déclaré
        /// intervalle au catalogue (ex. Stanford non déclaré = panneau masqué, aucune commande envoyée).</summary>
        public bool ShowIntervalleConfig =>
            MesureConfig.TypeMesure == TypeMesure.Interval
            && !MesureConfig.InitManu
            && ModeleCatalogueSelectionne()?.Parametres?.Intervalle?.Actif == true;

        /// <summary>Bloc des réglages "1 voie" (start et stop sur la voie 1).</summary>
        public bool ShowIntervUneVoie => ShowIntervalleConfig && !MesureConfig.IntervDeuxVoies;

        /// <summary>Bloc des réglages "2 voies" (start voie 1, stop voie 2).</summary>
        public bool ShowIntervDeuxVoies => ShowIntervalleConfig && MesureConfig.IntervDeuxVoies;

        /// <summary>Hold-off : utile uniquement en 1 voie (inhiber le 1er front montant→montant).</summary>
        public bool ShowHoldoff => ShowIntervUneVoie;

        // Sélecteur 1 voie / 2 voies
        public bool IntervModeUneVoie
        {
            get => !MesureConfig.IntervDeuxVoies;
            set { if (value) { MesureConfig.IntervDeuxVoies = false; NotifierIntervalle(); } }
        }
        public bool IntervModeDeuxVoies
        {
            get => MesureConfig.IntervDeuxVoies;
            set { if (value) { MesureConfig.IntervDeuxVoies = true; NotifierIntervalle(); } }
        }

        // Couplage voie 1
        public bool IntervDc1
        {
            get => MesureConfig.IntervDc1;
            set { MesureConfig.IntervDc1 = value; OnPropertyChanged(); OnPropertyChanged(nameof(IntervAc1)); }
        }
        public bool IntervAc1
        {
            get => !MesureConfig.IntervDc1;
            set { MesureConfig.IntervDc1 = !value; OnPropertyChanged(); OnPropertyChanged(nameof(IntervDc1)); }
        }

        // Impedance voie 1
        public bool IntervImp50_1
        {
            get => MesureConfig.IntervImp50_1;
            set { MesureConfig.IntervImp50_1 = value; OnPropertyChanged(); OnPropertyChanged(nameof(IntervImp1M_1)); }
        }
        public bool IntervImp1M_1
        {
            get => !MesureConfig.IntervImp50_1;
            set { MesureConfig.IntervImp50_1 = !value; OnPropertyChanged(); OnPropertyChanged(nameof(IntervImp50_1)); }
        }

        // Pente du départ
        public bool IntervStartMontant
        {
            get => MesureConfig.IntervStartMontant;
            set { MesureConfig.IntervStartMontant = value; OnPropertyChanged(); OnPropertyChanged(nameof(IntervStartDescendant)); }
        }
        public bool IntervStartDescendant
        {
            get => !MesureConfig.IntervStartMontant;
            set { MesureConfig.IntervStartMontant = !value; OnPropertyChanged(); OnPropertyChanged(nameof(IntervStartMontant)); }
        }

        // Pente de l'arrêt
        public bool IntervStopMontant
        {
            get => MesureConfig.IntervStopMontant;
            set { MesureConfig.IntervStopMontant = value; OnPropertyChanged(); OnPropertyChanged(nameof(IntervStopDescendant)); }
        }
        public bool IntervStopDescendant
        {
            get => !MesureConfig.IntervStopMontant;
            set { MesureConfig.IntervStopMontant = !value; OnPropertyChanged(); OnPropertyChanged(nameof(IntervStopMontant)); }
        }

        // Couplage voie 2 (mode 2 voies)
        public bool IntervDc2
        {
            get => MesureConfig.IntervDc2;
            set { MesureConfig.IntervDc2 = value; OnPropertyChanged(); OnPropertyChanged(nameof(IntervAc2)); }
        }
        public bool IntervAc2
        {
            get => !MesureConfig.IntervDc2;
            set { MesureConfig.IntervDc2 = !value; OnPropertyChanged(); OnPropertyChanged(nameof(IntervDc2)); }
        }

        // Impedance voie 2 (mode 2 voies)
        public bool IntervImp50_2
        {
            get => MesureConfig.IntervImp50_2;
            set { MesureConfig.IntervImp50_2 = value; OnPropertyChanged(); OnPropertyChanged(nameof(IntervImp1M_2)); }
        }
        public bool IntervImp1M_2
        {
            get => !MesureConfig.IntervImp50_2;
            set { MesureConfig.IntervImp50_2 = !value; OnPropertyChanged(); OnPropertyChanged(nameof(IntervImp50_2)); }
        }

        // Seuils (V) et hold-off (ns) : backing texte conservé tel quel (pas reformaté depuis le double
        // à chaque frappe, sinon "1.", "-", "0,5"... deviennent impossibles à saisir).
        // Le double est mis à jour uniquement quand parsable. SyncIntervalleTexte() resynchronise au chargement.
        private string _intervSeuilStartTexte = "1";
        private string _intervSeuilStopTexte = "1";
        private string _intervHoldoffTexte = "0";

        public string IntervSeuilStartTexte
        {
            get => _intervSeuilStartTexte;
            set
            {
                _intervSeuilStartTexte = value ?? string.Empty;
                if (double.TryParse(_intervSeuilStartTexte.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var v))
                    MesureConfig.IntervSeuilStart = v;
                OnPropertyChanged();
            }
        }
        public string IntervSeuilStopTexte
        {
            get => _intervSeuilStopTexte;
            set
            {
                _intervSeuilStopTexte = value ?? string.Empty;
                if (double.TryParse(_intervSeuilStopTexte.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var v))
                    MesureConfig.IntervSeuilStop = v;
                OnPropertyChanged();
            }
        }
        public string IntervHoldoffTexte
        {
            get => _intervHoldoffTexte;
            set
            {
                _intervHoldoffTexte = value ?? string.Empty;
                if (double.TryParse(_intervHoldoffTexte.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out var v))
                    MesureConfig.IntervHoldoffNs = v < 0 ? 0 : v;
                OnPropertyChanged();
            }
        }
        /// <summary>Resynchronise les champs texte intervalle depuis le modèle (au chargement d'une
        /// config). À ne PAS appeler pendant la frappe, sinon on écrase la saisie en cours.</summary>
        private void SyncIntervalleTexte()
        {
            _intervSeuilStartTexte = MesureConfig.IntervSeuilStart.ToString(CultureInfo.InvariantCulture);
            _intervSeuilStopTexte = MesureConfig.IntervSeuilStop.ToString(CultureInfo.InvariantCulture);
            _intervHoldoffTexte = MesureConfig.IntervHoldoffNs.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>Notifie l'UI de tout l'état du panneau intervalle (visibilités + valeurs).</summary>
        private void NotifierIntervalle()
        {
            OnPropertyChanged(nameof(ShowIntervalleConfig));
            OnPropertyChanged(nameof(ShowIntervUneVoie));
            OnPropertyChanged(nameof(ShowIntervDeuxVoies));
            OnPropertyChanged(nameof(ShowHoldoff));
            OnPropertyChanged(nameof(IntervModeUneVoie));
            OnPropertyChanged(nameof(IntervModeDeuxVoies));
            OnPropertyChanged(nameof(IntervDc1));
            OnPropertyChanged(nameof(IntervAc1));
            OnPropertyChanged(nameof(IntervImp50_1));
            OnPropertyChanged(nameof(IntervImp1M_1));
            OnPropertyChanged(nameof(IntervStartMontant));
            OnPropertyChanged(nameof(IntervStartDescendant));
            OnPropertyChanged(nameof(IntervStopMontant));
            OnPropertyChanged(nameof(IntervStopDescendant));
            OnPropertyChanged(nameof(IntervDc2));
            OnPropertyChanged(nameof(IntervAc2));
            OnPropertyChanged(nameof(IntervImp50_2));
            OnPropertyChanged(nameof(IntervImp1M_2));
            OnPropertyChanged(nameof(IntervSeuilStartTexte));
            OnPropertyChanged(nameof(IntervSeuilStopTexte));
            OnPropertyChanged(nameof(IntervHoldoffTexte));
        }

        /// <summary>Construit les commandes SCPI d'intervalle depuis les templates catalogue (rien en dur).
        /// Remplit {V}/{C}/{Z}/{S}/{P}/{T} avec les choix du panneau (1/2 voies, couplage, impédance, seuils, pentes, hold-off).</summary>
        private List<string> ConstruireCommandesIntervalle()
        {
            var c = new List<string>();
            var cfg = ModeleCatalogueSelectionne()?.Parametres?.Intervalle;
            if (cfg == null || !cfg.Actif) return c;   // appareil sans intervalle

            string Coup(bool dc) => dc ? "DC" : "AC";
            string Imp(bool z50) => z50 ? "50" : "1E6";
            string Pente(bool montant) => montant ? "POS" : "NEG";
            string V(double v) => v.ToString(CultureInfo.InvariantCulture);

            // Remplit un template avec les jetons fournis ; null si template vide.
            string? Fill(string tpl, int voie, string? coup = null, string? imp = null,
                         string? seuil = null, string? pente = null, string? temps = null)
            {
                if (string.IsNullOrWhiteSpace(tpl)) return null;
                string s = tpl.Replace("{V}", voie.ToString(CultureInfo.InvariantCulture));
                if (coup != null) s = s.Replace("{C}", coup);
                if (imp != null) s = s.Replace("{Z}", imp);
                if (seuil != null) s = s.Replace("{S}", seuil);
                if (pente != null) s = s.Replace("{P}", pente);
                if (temps != null) s = s.Replace("{T}", temps);
                return s;
            }
            void Add(string? cmd) { if (!string.IsNullOrWhiteSpace(cmd)) c.Add(cmd!); }

            if (!MesureConfig.IntervDeuxVoies)
            {
                // 1 voie : start/stop sur voie 1 (LEV1/SLOP1, LEV2/SLOP2)
                Add(cfg.Conf1Voie);
                Add(Fill(cfg.Couplage, 1, coup: Coup(MesureConfig.IntervDc1)));
                Add(Fill(cfg.Impedance, 1, imp: Imp(MesureConfig.IntervImp50_1)));
                Add(Fill(cfg.SeuilStart, 1, seuil: V(MesureConfig.IntervSeuilStart)));
                Add(Fill(cfg.PenteStart, 1, pente: Pente(MesureConfig.IntervStartMontant)));
                Add(Fill(cfg.SeuilStop1Voie, 1, seuil: V(MesureConfig.IntervSeuilStop)));
                Add(Fill(cfg.PenteStop1Voie, 1, pente: Pente(MesureConfig.IntervStopMontant)));

                // Hold-off : inhibition du 1er front. Template peut chaîner plusieurs commandes -> split.
                if (MesureConfig.IntervHoldoffNs > 0 && !string.IsNullOrWhiteSpace(cfg.Holdoff))
                {
                    string sec = string.Format(CultureInfo.InvariantCulture, "{0}E-9", V(MesureConfig.IntervHoldoffNs));
                    foreach (var cmd in SplitCommandesScpi(cfg.Holdoff.Replace("{T}", sec)))
                        Add(cmd);
                }
            }
            else
            {
                // 2 voies : start = voie 1, stop = voie 2 (via SeuilStart/PenteStart)
                Add(cfg.Conf2Voies);
                Add(Fill(cfg.Couplage, 1, coup: Coup(MesureConfig.IntervDc1)));
                Add(Fill(cfg.Impedance, 1, imp: Imp(MesureConfig.IntervImp50_1)));
                Add(Fill(cfg.SeuilStart, 1, seuil: V(MesureConfig.IntervSeuilStart)));
                Add(Fill(cfg.PenteStart, 1, pente: Pente(MesureConfig.IntervStartMontant)));
                Add(Fill(cfg.Couplage, 2, coup: Coup(MesureConfig.IntervDc2)));
                Add(Fill(cfg.Impedance, 2, imp: Imp(MesureConfig.IntervImp50_2)));
                Add(Fill(cfg.SeuilStart, 2, seuil: V(MesureConfig.IntervSeuilStop)));
                Add(Fill(cfg.PenteStart, 2, pente: Pente(MesureConfig.IntervStopMontant)));
            }

            return c;
        }

        /// <summary>Génère et envoie les commandes SCPI d'intervalle, puis les mémorise pour rejeu post-*RST.</summary>
        private async Task EnvoyerCommandesIntervalleAsync()
        {
            var commandes = ConstruireCommandesIntervalle();
            MesureConfig.CommandesScpiReglages = new List<string>(commandes);
            if (commandes.Count == 0) return;

            var modele = ModeleCatalogueSelectionne();
            var det = AppareilSelectionne?.Detecte;
            bool estFixe = AppareilSelectionne?.EstFixe == true;
            if (modele == null || (det == null && !estFixe))
            {
                // Appareil non joignable : config gardée localement, sera rejouée au lancement.
                JournalLog.Warn(CategorieLog.Configuration, "INTERVALLE_SCPI_LOCAL",
                    "Config intervalle conservée localement (appareil non joignable pour envoi immédiat).");
                return;
            }

            int board = det?.Board ?? 0;
            int adresse = det?.Adresse ?? AppareilSelectionne!.AdresseFixe;
            string adresseCourte = det?.AdresseCourte ?? $"GPIB{board}::{adresse}";

            JournalLog.Info(CategorieLog.Configuration, "INTERVALLE_SCPI_DEBUT",
                $"Envoi de {commandes.Count} commande(s) d'intervalle à {modele.Nom} @ {adresseCourte}.");

            try
            {
                using var driver = new VisaIeeeDriver(board);
                await AppareilScpiService.EnvoyerAsync(modele, adresse, commandes, driver);

                JournalLog.Info(CategorieLog.Configuration, "INTERVALLE_SCPI_OK",
                    $"Configuration d'intervalle envoyée à {modele.Nom}.");
            }
            catch (Exception ex)
            {
                JournalLog.Erreur(CategorieLog.Configuration, "INTERVALLE_SCPI_ERR",
                    $"Échec de l'envoi SCPI d'intervalle à {modele.Nom} : {ex.Message}",
                    new { ex.GetType().Name, Commandes = commandes });

                MessageBox.Show(
                    $"Erreur lors de l'envoi des commandes d'intervalle à {modele.Nom} :\n\n{ex.Message}\n\n"
                    + "Les réglages n'ont pas été appliqués sur l'appareil — la configuration locale "
                    + "est quand même conservée.",
                    "Erreur de communication GPIB",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // Source du signal : visible en type Fréquence uniquement.
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

        // Indirect : baie uniquement, pas en intervalle / tachy / stroboscope.
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
            NotifierIntervalle();
        }

        public void OnTypeMesureChanged()
        {
            if (MesureConfig.TypeMesure == TypeMesure.Interval)
                MesureConfig.NbMesures = 1;
            else
                MesureConfig.NbMesures = 30;

            // Hors Fréquence : source forcée sur Fréquencemètre par défaut.
            if (MesureConfig.TypeMesure != TypeMesure.Frequence)
                MesureConfig.SourceMesure = SourceMesure.Frequencemetre;

            // Mode indirect non dispo -> bascule sur Direct.
            if (MesureConfig.ModeMesure == ModeMesure.Indirect && !IndirectDisponible)
                MesureConfig.ModeMesure = ModeMesure.Direct;

            // Quitte Stabilité : reset GateIndices à 1 élément.
            // Sans ça, une mesure Fréquence après une Stab bouclerait sur toutes les gates sélectionnées.
            if (MesureConfig.TypeMesure != TypeMesure.Stabilite
                && MesureConfig.GateIndices.Count > 1)
            {
                MesureConfig.GateIndex = MesureConfig.GateIndices[0]; // setter remet la liste à 1 élément
            }

            // Intervalle -> force init manuelle (conforme F_Main.pas:987, Racal-Dana et Stanford).
            // Reste décochable par l'opérateur. Autres types sans init manuelle -> décoche.
            if (MesureConfig.TypeMesure == TypeMesure.Interval)
                MesureConfig.InitManu = true;
            else if (!InitManuDisponible && MesureConfig.InitManu)
                MesureConfig.InitManu = false;

            // Les modules d'incertitude dépendent de TypeMesure -> relister.
            RebuildModulesIncertitude();

            OnPropertyChanged(nameof(MesureConfig));
            RefreshAll();
            NotifierEtapes();   // changement type peut redéclencher AModulesIncertitude
        }

        public Action<bool>? CloseAction { get; set; }

        /// <summary>Valide la config : envoie les commandes SCPI puis ferme. Bloque sans N° FI.</summary>
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

            // Portage exact de F_Configuration.pas:104-127.

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
                // ASERi inaccessible : laisse le choix continuer sans vérif ou annuler.
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
                // FI inconnue : refus identique au Delphi
                MessageBox.Show(
                    $"La FI n° {sFI} n'existe pas dans ASERi !",
                    "Erreur",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            if (MesureConfig.InitManu)
            {
                // Init manuelle : aucune commande envoyée, réglages vidés (pas de rejeu post-*RST).
                MesureConfig.CommandesScpiReglages = new List<string>();
            }
            else if (MesureConfig.TypeMesure == TypeMesure.Interval)
            {
                // Intervalle piloté logiciel : panneau dédié -> SCPI CONF:TINT.
                await EnvoyerCommandesIntervalleAsync();
            }
            else
            {
                await EnvoyerCommandesScpiAsync();
            }

            // Intervalle + init manuelle : rappel de vérification avant lancement (aucune commande envoyée).
            if (MesureConfig.TypeMesure == TypeMesure.Interval && MesureConfig.InitManu)
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

        /// <summary>Commandes SCPI automatiques selon TypeMesure + VoieActive (cf. SelectionnerOptionAuto pour chaque réglage Auto).</summary>
        private List<string> CalculerCommandesAutomatiques(string idModele)
        {
            _ = idModele; // gardé pour signature future, mais on lit le catalogue directement
            var commandes = new List<string>();
            var modele = ModeleCatalogueSelectionne();
            if (modele == null) return commandes;

            // "Mode de mesure" inclus même sans marquage Auto : CONF:FREQ toujours calculé selon la voie active.
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

        /// <summary>Choisit l'option Auto par matching de libellés : "TIAB" si Intervalle, "Voie A/B/C" selon
        /// la voie active, "CONT" en Stabilité, "AUTO" sinon ; fallback : 1ère option du catalogue.</summary>
        private OptionReglage? SelectionnerOptionAuto(ReglageAppareil reglage)
        {
            if (reglage.Options.Count == 0) return null;

            // Mode de mesure : Intervalle en priorité (TIAB), sinon par voie active.
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

            // Résolution : CONT si Stabilité, AUTO sinon.
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

            // Fallback : 1ère option du catalogue.
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

            // Adresse cible : appareil détecté (scan) ou adresse fixe (legacy sans *IDN?).
            // Sans cette branche, le mode adresses fixes laissait CommandesScpiReglages vide -> impédance/couplage non rejoués.
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

            // Réglages de la voie active + le mode global (n'affecte pas les voies inutilisées).
            var reglagesApplicables = new List<ReglageDynamiqueViewModel>();
            reglagesApplicables.AddRange(MesureConfig.VoieActive switch
            {
                VoieActive.A => ReglagesVoieA,
                VoieActive.B => ReglagesVoieB,
                VoieActive.C => ReglagesVoieC,
                _ => ReglagesVoieA
            });
            reglagesApplicables.AddRange(ReglagesMode);

            // Split les commandes chaînées ";:" en writes GPIB distincts :
            // le firmware 53230A tronque parfois les chaînes concaténées et ignore silencieusement la suite.
            var commandesUtilisateur = reglagesApplicables
                .Select(r => r.CommandeSelectionnee)
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .SelectMany(c => SplitCommandesScpi(c!))
                .ToList();

            // Commandes auto (ex. :SENS:FREQ:MODE) selon TypeMesure + VoieActive.
            var commandesAuto = CalculerCommandesAutomatiques(modele.Id);

            // ORDRE CRITIQUE (53230A) : CONF:FREQ/:FUNC réinitialise tous les INP:* (LEV:AUTO ON, LEV 0V).
            // Toute commande INP:* avant CONF:* est écrasée.
            // Ordre : 1. CONF:FREQ/:FUNC  2. :SENS:FREQ:MODE  3. INP:IMP/COUP/FILT  4. LEV:AUTO OFF + LEV
            bool EstCommandeSetup(string cmd) =>
                cmd.StartsWith("CONF:", StringComparison.OrdinalIgnoreCase) ||
                cmd.StartsWith(":FUNC", StringComparison.OrdinalIgnoreCase) ||
                cmd.StartsWith(":CONF", StringComparison.OrdinalIgnoreCase);

            // CONF:FREQ/:FUNC vient des commandes auto : on l'extrait de toutes les sources pour le mettre en tête.
            var commandesSetup = commandesUtilisateur.Where(EstCommandeSetup)
                .Concat(commandesAuto.Where(EstCommandeSetup)).ToList();
            var commandesAutoReste = commandesAuto.Where(c => !EstCommandeSetup(c)).ToList();
            var commandesEntree = commandesUtilisateur.Where(c => !EstCommandeSetup(c)).ToList();

            var commandes = new List<string>();
            commandes.AddRange(commandesSetup);       // 1. CONF:FREQ / :FUNC (réinitialise les inputs)
            commandes.AddRange(commandesAutoReste);   // 2. :SENS:FREQ:MODE (résolution)
            commandes.AddRange(commandesEntree);      // 3. INP1:IMP/COUP/FILT/RANG + LEV:AUTO OFF + LEV

            // Mémorisées dans Mesure : l'orchestrator les rejoue après *RST pendant la boucle.
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

        /// <summary>Découpe une chaîne SCPI ";:" en commandes individuelles préfixées ":" (workaround firmware 53230A).</summary>
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
                // 1ère partie : garde son ":" initial. Suivantes : ":" perdu dans le split, on le remet.
                yield return i == 0 ? p : ":" + p;
            }
        }
    }
}
