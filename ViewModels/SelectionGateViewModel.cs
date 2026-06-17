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
    /// ViewModel de l'écran de choix du temps de porte.
    /// <para>
    /// Hors Stabilité, on présente une combo des gates disponibles et on ne renvoie qu'un
    /// seul indice dans <see cref="IndicesGatesResultats"/> (la liste ne contient alors qu'un élément).
    /// </para>
    /// <para>
    /// En Stabilité, les gates s'affichent sous forme de cases à cocher : l'utilisateur peut
    /// en sélectionner plusieurs pour les balayer l'une après l'autre — c'est la version
    /// moderne des anciennes « procédures auto ». Les combinaisons qui reviennent souvent sont
    /// conservées comme presets dans <see cref="PresetsStabiliteService"/>, modifiables sans
    /// avoir à recompiler.
    /// </para>
    /// </summary>
    public partial class SelectionGateViewModel : ObservableObject
    {
        private readonly Mesure _mesure;

        [ObservableProperty] private TypeMesure _typeMesure;

        /// <summary>Indice retenu en mode "gate fixe" (hors Stabilité). -1 signifie : rien de sélectionné.</summary>
        [ObservableProperty] private int _gateSelectionneIndex = -1;

        /// <summary>Les gates disponibles, présentées en cases à cocher pour le mode Stabilité.</summary>
        public ObservableCollection<GateCochable> GatesDisponibles { get; } = new();

        /// <summary>Les presets de balayage, lus depuis <see cref="PresetsStabiliteService"/>.</summary>
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

        /// <summary>Le résultat renvoyé : un seul indice pour une mesure simple, plusieurs pour la stabilité.</summary>
        public List<int> IndicesGatesResultats { get; private set; } = new();

        public Action<bool>? CloseAction { get; set; }

        public SelectionGateViewModel(Mesure mesure)
        {
            _mesure = mesure;
            TypeMesure = mesure.TypeMesure;

            ConstruireGatesDisponibles();

            if (IsStabilite)
            {
                // On présélectionne le 1er preset (l'équivalent de l'ancienne procédure Auto) pour
                // que l'écran s'ouvre déjà sur un choix cohérent, sans que l'utilisateur ait à cocher.
                PresetSelectionne = Presets.FirstOrDefault();
                if (PresetSelectionne != null) AppliquerPreset(PresetSelectionne);
            }
            else
            {
                // Valeur par défaut, reprise du comportement d'origine : 10 s pour FreqAvantInterv,
                // 1 s dans les autres cas. On cherche l'index canonique correspondant dans GatesDisponibles.
                int defaut = TypeMesure == TypeMesure.FreqAvantInterv ? 9 : 6;
                GateSelectionneIndex = TrouverPositionListePourSlot(defaut);
            }
        }

        /// <summary>
        /// Monte la liste des gates à afficher. On la limite à celles cochées dans le catalogue
        /// pour l'appareil sélectionné (voir <see cref="ConfigurationViewModel.GateTimes"/>). Et si
        /// aucun appareil du catalogue ne correspond, on retombe sur l'échelle canonique complète.
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
        /// Traduit un slot canonique (0..15) en sa position dans la liste affichée
        /// <see cref="GatesDisponibles"/>, qui n'en contient parfois qu'une partie. Renvoie -1
        /// quand le slot n'est pas proposé pour l'appareil courant.
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

        /// <summary>Coche les cases dont le libellé figure dans le preset et décoche tout le reste.</summary>
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

                // On journalise ce qui part réellement vers l'orchestrator. Pratique pour traquer
                // les désynchros entre l'UI et le modèle (typiquement une case décochée à l'écran
                // alors que EstCoche est resté à true).
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
