using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Metrologo.Models;

namespace Metrologo.Services.Catalogue
{
    /// <summary>
    /// Catalogue local des modèles d'appareils enregistrés par l'utilisateur.
    /// Persisté en JSON dans <c>%LocalAppData%\Metrologo\AppareilsCatalogue.json</c>.
    ///
    /// Singleton accessible via <see cref="Instance"/>. Migration vers un fichier partagé
    /// prévue plus tard (il suffira de changer le chemin).
    /// </summary>
    public class CatalogueAppareilsService
    {
        private static readonly Lazy<CatalogueAppareilsService> _instance = new(() => new());
        public static CatalogueAppareilsService Instance => _instance.Value;

        private readonly string _cheminFichier;

        /// <summary>Collection observable des modèles — à binder depuis l'UI admin.</summary>
        public ObservableCollection<ModeleAppareil> Modeles { get; } = new();

        /// <summary>Notifié après tout ajout / modification / suppression / rechargement.</summary>
        public event EventHandler? CatalogueChange;

        private CatalogueAppareilsService()
        {
            _cheminFichier = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Metrologo", "AppareilsCatalogue.json");
        }

        public string CheminFichier => _cheminFichier;

        public async Task ChargerAsync()
        {
            try
            {
                if (!File.Exists(_cheminFichier)) { Modeles.Clear(); NotifierChange(); return; }

                var json = await File.ReadAllTextAsync(_cheminFichier);
                var parsed = JsonSerializer.Deserialize<List<ModeleAppareil>>(json)
                             ?? new List<ModeleAppareil>();

                Modeles.Clear();
                foreach (var m in parsed) Modeles.Add(m);
                NotifierChange();
            }
            catch (Exception)
            {
                // Fichier corrompu : on repart sur un catalogue vide plutôt que de planter.
                Modeles.Clear();
                NotifierChange();
            }
        }

        public async Task SauvegarderAsync()
        {
            var dir = Path.GetDirectoryName(_cheminFichier);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

            var json = JsonSerializer.Serialize(Modeles.ToList(),
                new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(_cheminFichier, json);
        }

        /// <summary>Ajoute un modèle. Tout utilisateur peut le faire.</summary>
        public async Task AjouterAsync(ModeleAppareil modele)
        {
            if (string.IsNullOrEmpty(modele.Id)) modele.Id = GenererId(modele.Nom);
            Modeles.Add(modele);
            await SauvegarderAsync();
            NotifierChange();
        }

        /// <summary>Modifie un modèle existant. Réservé à l'admin (garde côté UI).</summary>
        public async Task ModifierAsync(string id, Action<ModeleAppareil> modification)
        {
            var cible = Modeles.FirstOrDefault(m => m.Id == id);
            if (cible == null) return;
            modification(cible);
            await SauvegarderAsync();
            NotifierChange();
        }

        /// <summary>Supprime un modèle. Réservé à l'admin (garde côté UI).</summary>
        public async Task SupprimerAsync(string id)
        {
            var cible = Modeles.FirstOrDefault(m => m.Id == id);
            if (cible == null) return;
            Modeles.Remove(cible);
            await SauvegarderAsync();
            NotifierChange();
        }

        /// <summary>Cherche un modèle correspondant à la signature IDN donnée.</summary>
        public ModeleAppareil? TrouverParIdn(string? fabricant, string? modele)
            => Modeles.FirstOrDefault(m => m.Correspond(fabricant, modele));

        /// <summary>Vrai si un modèle correspondant à cet IDN est déjà enregistré.</summary>
        public bool EstDansCatalogue(string? fabricant, string? modele)
            => TrouverParIdn(fabricant, modele) != null;

        // ---------------- Interne ----------------

        private void NotifierChange() => CatalogueChange?.Invoke(this, EventArgs.Empty);

        private static string GenererId(string nom)
        {
            string slug = new string((nom ?? "modele")
                .ToLowerInvariant()
                .Select(c => char.IsLetterOrDigit(c) ? c : '-')
                .ToArray())
                .Trim('-');
            string suffixe = Guid.NewGuid().ToString("N").Substring(0, 6);
            return string.IsNullOrEmpty(slug) ? $"modele-{suffixe}" : $"{slug}-{suffixe}";
        }
    }
}
