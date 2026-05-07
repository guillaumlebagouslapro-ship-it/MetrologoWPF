using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Metrologo.Models;
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
            }
        }

        /// <summary>Vrai si au moins un module couvre le type de mesure courant.</summary>
        public bool AModulesIncertitude => ModulesDisponibles.Count > 0;

        /// <summary>
        /// Reconstruit la liste des modules selon <see cref="MesureConfig.TypeMesure"/>.
        /// Lit le sous-dossier dédié à ce type (ex. <c>Incertitudes\Frequence\</c>).
        /// Restaure la sélection persistée (<c>NumModuleIncertitude</c>) si elle est encore valide.
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

            OnPropertyChanged(nameof(ModuleSelectionne));
            OnPropertyChanged(nameof(AModulesIncertitude));
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

            await EnvoyerCommandesScpiAsync();
            CloseAction?.Invoke(true);
        }

        [RelayCommand] private void Annuler() => CloseAction?.Invoke(false);

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

            var commandes = reglagesApplicables
                .Select(r => r.CommandeSelectionnee)
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Select(c => c!)
                .ToList();

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
    }
}
