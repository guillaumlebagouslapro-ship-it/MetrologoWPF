using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using Metrologo.Models;

namespace Metrologo.Services
{
    /// <summary>
    /// Catalogue des presets de balayage des mesures de stabilité, en JSON
    /// (<see cref="CheminsMetrologo.FichierPresetsStabilite"/>). L'utilisateur crée ses propres
    /// presets depuis la fenêtre de stabilité ; au premier démarrage, quelques presets par défaut
    /// (dont les deux procédures Auto historiques) sont semés.
    /// </summary>
    public class PresetsStabiliteService
    {
        private static readonly Lazy<PresetsStabiliteService> _instance = new(() => new());
        public static PresetsStabiliteService Instance => _instance.Value;

        private readonly string _cheminFichier;

        public ObservableCollection<PresetStabilite> Presets { get; } = new();

        public event EventHandler? PresetsChange;

        private PresetsStabiliteService()
        {
            _cheminFichier = CheminsMetrologo.FichierPresetsStabilite;
        }

        public string CheminFichier => _cheminFichier;

        public async Task ChargerAsync()
        {
            try
            {
                if (!File.Exists(_cheminFichier))
                {
                    Presets.Clear();
                    foreach (var p in PresetsParDefaut()) Presets.Add(p);
                    await SauvegarderAsync();
                    NotifierChange();
                    return;
                }

                var json = await File.ReadAllTextAsync(_cheminFichier);
                var parsed = JsonSerializer.Deserialize<List<PresetStabilite>>(json)
                             ?? new List<PresetStabilite>();

                Presets.Clear();
                foreach (var p in parsed) Presets.Add(p);
                NotifierChange();
            }
            catch
            {
                // Fichier corrompu : presets par défaut plutôt qu'un plantage.
                Presets.Clear();
                foreach (var p in PresetsParDefaut()) Presets.Add(p);
                NotifierChange();
            }
        }

        public async Task SauvegarderAsync()
        {
            var dir = Path.GetDirectoryName(_cheminFichier);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            var json = JsonSerializer.Serialize(Presets.ToList(),
                new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(_cheminFichier, json);
        }

        public async Task AjouterOuMettreAJourAsync(PresetStabilite preset)
        {
            var existant = Presets.FirstOrDefault(p =>
                string.Equals(p.Nom, preset.Nom, StringComparison.OrdinalIgnoreCase));

            if (existant != null)
            {
                existant.GatesLibelles = new List<string>(preset.GatesLibelles);
            }
            else
            {
                Presets.Add(preset);
            }

            await SauvegarderAsync();
            NotifierChange();
        }

        public async Task SupprimerAsync(string nom)
        {
            var cible = Presets.FirstOrDefault(p =>
                string.Equals(p.Nom, nom, StringComparison.OrdinalIgnoreCase));
            if (cible == null) return;

            Presets.Remove(cible);
            await SauvegarderAsync();
            NotifierChange();
        }

        private void NotifierChange() => PresetsChange?.Invoke(this, EventArgs.Empty);

        /// <summary>
        /// Presets semés au premier démarrage : les deux procédures auto historiques
        /// (10 ms → 100 s et 10 ms → 10 s) et un balayage court. Modifiables et supprimables.
        /// </summary>
        private static IEnumerable<PresetStabilite> PresetsParDefaut() => new[]
        {
            new PresetStabilite
            {
                Nom = "Balayage standard (10 ms → 100 s)",
                GatesLibelles = new List<string>
                {
                    "10 ms", "20 ms", "50 ms", "100 ms", "200 ms", "500 ms",
                    "1 s", "2 s", "5 s", "10 s", "20 s", "50 s", "100 s"
                }
            },
            new PresetStabilite
            {
                Nom = "Balayage rapide (10 ms → 10 s)",
                GatesLibelles = new List<string>
                {
                    "10 ms", "20 ms", "50 ms", "100 ms", "200 ms", "500 ms",
                    "1 s", "2 s", "5 s", "10 s"
                }
            },
            new PresetStabilite
            {
                Nom = "Contrôle court (10 ms, 100 ms, 1 s)",
                GatesLibelles = new List<string> { "10 ms", "100 ms", "1 s" }
            }
        };
    }
}
