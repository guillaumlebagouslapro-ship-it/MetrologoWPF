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
        }

        /// <summary>
        /// Liste unifiée des appareils disponibles dans la dropdown :
        /// les 3 types catalogue (Stanford/Racal/EIP) + tout appareil détecté inconnu du catalogue.
        /// Chaque type catalogue indique s'il est actuellement branché ou non.
        /// </summary>
        public ObservableCollection<OptionAppareil> Appareils { get; } = new();

        private void RebuildAppareils()
        {
            // Mémoriser la sélection courante pour la retrouver après reconstruction de la liste.
            var selType = _appareilSelectionne?.Type;
            var selIdnFab = _appareilSelectionne?.Detecte?.Fabricant;
            var selIdnMod = _appareilSelectionne?.Detecte?.Modele;

            // Re-synchronise ModeleReconnu pour chaque appareil détecté (le catalogue peut avoir changé).
            foreach (var det in EtatApplication.AppareilsDetectes)
            {
                det.ModeleReconnu = CatalogueAppareilsService.Instance.TrouverParIdn(det.Fabricant, det.Modele);
            }

            Appareils.Clear();

            // 1) Types catalogue — toujours présents, marqués "détecté" si on les voit sur le bus
            foreach (TypeAppareilIEEE type in Enum.GetValues(typeof(TypeAppareilIEEE)))
            {
                var det = EtatApplication.AppareilsDetectes.FirstOrDefault(a => a.TypeReconnu == type);
                string nom = NomCatalogue(type);
                string suffixe = det != null ? $" — {det.AdresseCourte} ✓" : "  (non connecté)";
                Appareils.Add(new OptionAppareil
                {
                    Libelle = nom + suffixe,
                    Type = type,
                    Detecte = det
                });
            }

            // 2) Appareils détectés qui ne correspondent à aucun des 3 types catalogue historiques.
            //    Si présents au catalogue local → "catalogue", sinon "hors catalogue" (à enregistrer).
            foreach (var det in EtatApplication.AppareilsDetectes.Where(a => a.TypeReconnu == null))
            {
                string suffixe = det.ModeleReconnu != null
                    ? $" ✓  (catalogue)"
                    : $" ✓  (hors catalogue)";

                Appareils.Add(new OptionAppareil
                {
                    Libelle = $"{det.Libelle}{suffixe}",
                    Type = null,
                    Detecte = det
                });
            }

            // 3) Retrouver la sélection courante dans la nouvelle liste (par Type pour les 3 types
            //    historiques, sinon par IDN fabricant+modèle pour les appareils catalogue).
            if (selType.HasValue)
            {
                _appareilSelectionne = Appareils.FirstOrDefault(o => o.Type == selType);
            }
            else if (selIdnFab != null || selIdnMod != null)
            {
                _appareilSelectionne = Appareils.FirstOrDefault(o =>
                    o.Detecte != null
                    && o.Detecte.Fabricant == selIdnFab
                    && o.Detecte.Modele == selIdnMod);
            }

            OnPropertyChanged(nameof(AppareilSelectionne));
            RebuildReglagesDynamiques();
            RefreshAll();
        }

        private static string NomCatalogue(TypeAppareilIEEE t) => t switch
        {
            TypeAppareilIEEE.Stanford => "Stanford SR620",
            TypeAppareilIEEE.Racal    => "Racal-Dana 1996",
            TypeAppareilIEEE.EIP      => "EIP 545",
            _ => t.ToString()
        };

        private OptionAppareil? _appareilSelectionne;

        /// <summary>
        /// Sélection courante dans la ComboBox. Stockée explicitement pour supporter les appareils
        /// hors des 3 types historiques (Type == null) — on ne peut pas se contenter de dériver
        /// la sélection depuis <c>MesureConfig.Frequencemetre</c>.
        /// </summary>
        public OptionAppareil? AppareilSelectionne
        {
            get
            {
                // Si on a une sélection explicite et qu'elle est toujours dans la liste, on la garde.
                if (_appareilSelectionne != null && Appareils.Contains(_appareilSelectionne))
                    return _appareilSelectionne;

                // Sinon fallback sur l'option catalogue correspondant au Frequencemetre courant.
                return Appareils.FirstOrDefault(o => o.Type == MesureConfig.Frequencemetre)
                       ?? Appareils.FirstOrDefault();
            }
            set
            {
                if (ReferenceEquals(_appareilSelectionne, value)) return;
                _appareilSelectionne = value;

                if (value?.Type.HasValue == true)
                {
                    MesureConfig.Frequencemetre = value.Type.Value;
                    // Sélection d'un des 3 types historiques : pas de modèle catalogue explicite,
                    // l'orchestrator utilisera Metrologo.ini.
                    MesureConfig.IdModeleCatalogue = string.Empty;
                }
                else
                {
                    // Appareil hors des 3 types historiques → on vise le modèle catalogue si présent,
                    // sinon rien (l'orchestrator retombera en erreur explicite au lancement).
                    var modele = ModeleCatalogueSelectionne();
                    MesureConfig.IdModeleCatalogue = modele?.Id ?? string.Empty;
                }

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
                foreach (var reglage in modele.Reglages)
                {
                    var vm = new ReglageDynamiqueViewModel(reglage);
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

        public List<string> StanfordRanges => new() { "1 MΩ (A)", "50 Ω (A)", "UHF (C)" };
        public List<string> RacalRanges => new() { "Entrée A", "Entrée B", "Entrée C" };
        public List<string> EipRanges => new() { "Bande 1", "Bande 2", "Bande 3" };
        public List<string> Couplings => new() { "AC", "DC" };
        public List<string> GateTimes => new()
        {
            "10 ms", "20 ms", "50 ms",
            "100 ms", "200 ms", "500 ms",
            "1 s", "2 s", "5 s",
            "10 s", "20 s", "50 s", "100 s"
        };

        public List<int> MeasurementCounts => Enumerable.Range(1, 100).ToList();

        public bool IsStanford => MesureConfig.Frequencemetre == TypeAppareilIEEE.Stanford;
        public bool IsRacal => MesureConfig.Frequencemetre == TypeAppareilIEEE.Racal;
        public bool IsEip => MesureConfig.Frequencemetre == TypeAppareilIEEE.EIP;

        /// <summary>Vrai si l'appareil sélectionné est dans le catalogue local (hors 3 types historiques).</summary>
        public bool EstAppareilCatalogue => AppareilSelectionne?.Detecte?.ModeleReconnu != null
                                            || (AppareilSelectionne?.Detecte != null
                                                && CatalogueAppareilsService.Instance.EstDansCatalogue(
                                                    AppareilSelectionne.Detecte.Fabricant,
                                                    AppareilSelectionne.Detecte.Modele));

        // Les options hardcodées (GAMME D'ENTRÉE, COUPLAGE legacy) ne s'affichent QUE pour un
        // des 3 types historiques ET seulement si l'appareil n'a pas ses propres réglages dans le catalogue.
        public bool AfficherGammeStanford => IsStanford && !EstAppareilCatalogue;
        public bool AfficherGammeRacal    => IsRacal    && !EstAppareilCatalogue;
        public bool AfficherGammeEip      => IsEip      && !EstAppareilCatalogue;

        public bool ShowGateSettings =>
            MesureConfig.TypeMesure == TypeMesure.Frequence ||
            MesureConfig.TypeMesure == TypeMesure.Stabilite;

        public bool ShowCoupling => ((IsStanford && MesureConfig.InputIndex != 2) ||
                                     (IsRacal && MesureConfig.InputIndex != 2))
                                    && !EstAppareilCatalogue;

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

        // Indirect disponible : pas en paillasse, pas EIP, pas intervalle / tachy / stroboscope
        public bool IndirectDisponible =>
            EstSurBaie
            && !IsEip
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
            OnPropertyChanged(nameof(IsStanford));
            OnPropertyChanged(nameof(IsRacal));
            OnPropertyChanged(nameof(IsEip));
            OnPropertyChanged(nameof(EstAppareilCatalogue));
            OnPropertyChanged(nameof(AfficherGammeStanford));
            OnPropertyChanged(nameof(AfficherGammeRacal));
            OnPropertyChanged(nameof(AfficherGammeEip));
            OnPropertyChanged(nameof(IsModeDirect));
            OnPropertyChanged(nameof(IsModeIndirect));
            OnPropertyChanged(nameof(ShowGateSettings));
            OnPropertyChanged(nameof(ShowCoupling));
            OnPropertyChanged(nameof(ShowSourceMesure));
            OnPropertyChanged(nameof(IsSourceFrequencemetre));
            OnPropertyChanged(nameof(IsSourceGenerateur));
            OnPropertyChanged(nameof(IndirectDisponible));
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
