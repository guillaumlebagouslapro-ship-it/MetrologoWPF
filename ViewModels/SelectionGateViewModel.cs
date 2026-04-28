using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services;
using Metrologo.Services.Catalogue;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// ViewModel de la sélection du temps de porte.
    /// <para>
    /// Pour les types non-Stabilité, expose une combo des gates disponibles et renvoie
    /// un unique indice dans <see cref="IndicesGatesResultats"/> (1 élément).
    /// </para>
    /// <para>
    /// Pour la Stabilité, expose les gates disponibles en cases à cocher : l'utilisateur
    /// peut en cocher plusieurs pour balayer la liste séquentiellement (équivalent moderne
    /// des « procédures auto » historiques). Les combinaisons fréquentes sont stockées
    /// comme presets dans <see cref="PresetsStabiliteService"/>, éditables sans recompiler.
    /// </para>
    /// </summary>
    public partial class SelectionGateViewModel : ObservableObject
    {
        private readonly Mesure _mesure;

        [ObservableProperty] private TypeMesure _typeMesure;

        /// <summary>Indice sélectionné en mode "gate fixe" (types non-Stabilité). -1 = aucun.</summary>
        [ObservableProperty] private int _gateSelectionneIndex = -1;

        /// <summary>Cases à cocher des gates disponibles pour le mode Stabilité.</summary>
        public ObservableCollection<GateCochable> GatesDisponibles { get; } = new();

        /// <summary>Presets de balayage chargés depuis <see cref="PresetsStabiliteService"/>.</summary>
        public ObservableCollection<PresetStabilite> Presets => PresetsStabiliteService.Instance.Presets;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(PresetSelectionneNomCourt))]
        private PresetStabilite? _presetSelectionne;

        [ObservableProperty] private string _nomNouveauPreset = string.Empty;
        [ObservableProperty] private string _messageErreur = string.Empty;

        public bool IsStabilite => TypeMesure == TypeMesure.Stabilite;
        public string Titre => IsStabilite ? "Mesure de stabilité — choix du balayage" : "Sélection du temps de porte";
        public string SousTitre => IsStabilite
            ? "Cochez les temps de porte à balayer ou choisissez un preset"
            : "Définissez le temps de porte de la mesure";
        public bool HasError => !string.IsNullOrEmpty(MessageErreur);
        public string PresetSelectionneNomCourt => PresetSelectionne?.Nom ?? "(aucun)";

        /// <summary>Résultat final : 1 indice pour les mesures simples, N indices pour la stabilité.</summary>
        public List<int> IndicesGatesResultats { get; private set; } = new();

        public Action<bool>? CloseAction { get; set; }

        public SelectionGateViewModel(Mesure mesure)
        {
            _mesure = mesure;
            TypeMesure = mesure.TypeMesure;

            ConstruireGatesDisponibles();

            if (IsStabilite)
            {
                // Pré-sélectionne le 1er preset (≈ procédure Auto historique) pour que l'utilisateur
                // ait quelque chose de cohérent dès l'ouverture, sans devoir cocher manuellement.
                PresetSelectionne = Presets.FirstOrDefault();
                if (PresetSelectionne != null) AppliquerPreset(PresetSelectionne);
            }
            else
            {
                // Valeur par défaut conforme à l'ancien comportement : 10 s pour FreqAvantInterv,
                // 1 s sinon. On vise l'index canonique correspondant dans GatesDisponibles.
                int defaut = TypeMesure == TypeMesure.FreqAvantInterv ? 9 : 6;
                GateSelectionneIndex = TrouverPositionListePourSlot(defaut);
            }
        }

        /// <summary>
        /// Construit la liste des gates affichées : restreinte à celles cochées dans le catalogue
        /// pour l'appareil sélectionné (cf. <see cref="ConfigurationViewModel.GateTimes"/>). Si
        /// aucun appareil catalogue ne correspond, on tombe sur l'échelle canonique complète.
        /// </summary>
        private void ConstruireGatesDisponibles()
        {
            GatesDisponibles.Clear();

            var libelles = LibellesGatesPourMesure();
            foreach (var lib in libelles)
            {
                int slot = EnTetesMesureHelper.IndexDepuisLibelle(lib);
                if (slot < 0) continue;
                GatesDisponibles.Add(new GateCochable(lib) { SlotCanonique = slot });
            }
        }

        private IEnumerable<string> LibellesGatesPourMesure()
        {
            if (!string.IsNullOrEmpty(_mesure.IdModeleCatalogue))
            {
                var modele = CatalogueAppareilsService.Instance.Modeles
                    .FirstOrDefault(m => m.Id == _mesure.IdModeleCatalogue);
                if (modele != null && modele.Gates.Count > 0) return modele.Gates;
            }
            return EnTetesMesureHelper.LibellesCanoniques;
        }

        /// <summary>
        /// Convertit un slot canonique (0..15) en position dans la liste affichée
        /// <see cref="GatesDisponibles"/> (qui peut être un sous-ensemble). -1 si le slot
        /// n'est pas exposé pour l'appareil courant.
        /// </summary>
        private int TrouverPositionListePourSlot(int slotCanonique)
        {
            for (int i = 0; i < GatesDisponibles.Count; i++)
                if (GatesDisponibles[i].SlotCanonique == slotCanonique) return i;
            return GatesDisponibles.Count > 0 ? 0 : -1;
        }

        partial void OnPresetSelectionneChanged(PresetStabilite? value)
        {
            if (value != null && IsStabilite) AppliquerPreset(value);
        }

        /// <summary>Coche les cases qui correspondent aux libellés du preset, décoche les autres.</summary>
        private void AppliquerPreset(PresetStabilite preset)
        {
            var attendus = new HashSet<string>(
                preset.GatesLibelles.Select(NormaliserLibelle), StringComparer.OrdinalIgnoreCase);

            foreach (var g in GatesDisponibles)
                g.EstCoche = attendus.Contains(NormaliserLibelle(g.Libelle));
        }

        private static string NormaliserLibelle(string libelle)
            => string.Concat(libelle.Where(c => !char.IsWhiteSpace(c))).ToLowerInvariant();

        // -------------------- Commandes --------------------

        [RelayCommand]
        private async Task SauverCommePresetAsync()
        {
            if (!IsStabilite) return;

            var libellesCoches = GatesDisponibles.Where(g => g.EstCoche).Select(g => g.Libelle).ToList();
            if (libellesCoches.Count == 0)
            {
                MessageErreur = "Cochez au moins une gate avant de sauver le preset.";
                OnPropertyChanged(nameof(HasError));
                return;
            }

            string nom = (NomNouveauPreset ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(nom))
            {
                MessageErreur = "Saisissez un nom pour le preset.";
                OnPropertyChanged(nameof(HasError));
                return;
            }

            await PresetsStabiliteService.Instance.AjouterOuMettreAJourAsync(new PresetStabilite
            {
                Nom = nom,
                GatesLibelles = libellesCoches
            });

            PresetSelectionne = Presets.FirstOrDefault(p =>
                string.Equals(p.Nom, nom, StringComparison.OrdinalIgnoreCase));
            NomNouveauPreset = string.Empty;
            MessageErreur = string.Empty;
            OnPropertyChanged(nameof(HasError));
        }

        [RelayCommand]
        private async Task SupprimerPresetAsync()
        {
            if (!IsStabilite || PresetSelectionne == null) return;

            var confirm = MessageBox.Show(
                $"Supprimer le preset « {PresetSelectionne.Nom} » ?",
                "Confirmer", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes) return;

            string nom = PresetSelectionne.Nom;
            await PresetsStabiliteService.Instance.SupprimerAsync(nom);
            PresetSelectionne = Presets.FirstOrDefault();
        }

        [RelayCommand]
        private void Valider()
        {
            if (IsStabilite)
            {
                var indices = GatesDisponibles
                    .Where(g => g.EstCoche)
                    .Select(g => g.SlotCanonique)
                    .Distinct()
                    .OrderBy(i => i)
                    .ToList();

                // Trace ce qui est réellement transmis à l'orchestrator — utile pour diagnostiquer
                // les désynchros UI/modèle (ex: cases visuelles décochées mais EstCoche encore à true).
                int nbCochees = GatesDisponibles.Count(g => g.EstCoche);
                Services.Journal.Journal.Info(Services.Journal.CategorieLog.Configuration,
                    "STAB_GATES_VALIDEES",
                    $"Stab : {nbCochees}/{GatesDisponibles.Count} gates cochées → indices retenus : "
                    + $"{string.Join(",", indices)} (libellés : "
                    + $"{string.Join(",", GatesDisponibles.Where(g => g.EstCoche).Select(g => g.Libelle))})");

                if (indices.Count == 0)
                {
                    MessageErreur = "Cochez au moins un temps de porte à balayer.";
                    OnPropertyChanged(nameof(HasError));
                    return;
                }

                IndicesGatesResultats = indices;
            }
            else
            {
                if (GateSelectionneIndex < 0 || GateSelectionneIndex >= GatesDisponibles.Count)
                {
                    MessageErreur = "Sélectionnez un temps de porte.";
                    OnPropertyChanged(nameof(HasError));
                    return;
                }
                IndicesGatesResultats = new List<int> { GatesDisponibles[GateSelectionneIndex].SlotCanonique };
            }

            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);
    }
}
