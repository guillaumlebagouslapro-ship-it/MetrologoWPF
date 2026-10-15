using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using Metrologo.Models;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Catalogue
{
    /// <summary>
    /// Catalogue central des modèles d'appareils SCPI/IEEE : un unique JSON sur le partage réseau
    /// (<see cref="CheminsMetrologo.FichierCatalogueAppareils"/>), relu à chaque ChargerAsync et
    /// réécrit atomiquement après chaque modification. Si le fichier réseau est absent, l'ancien JSON
    /// hérité de l'époque SQL Server (AppareilsCatalogue.json) est repris. Multi-postes : le
    /// SemaphoreSlim et l'écriture atomique évitent la corruption, mais le dernier qui écrit
    /// l'emporte (un vrai merge demanderait un fichier .lock).
    /// </summary>
    public class CatalogueAppareilsService
    {
        private static readonly Lazy<CatalogueAppareilsService> _instance = new(() => new());
        public static CatalogueAppareilsService Instance => _instance.Value;

        public ObservableCollection<ModeleAppareil> Modeles { get; } = new();
        public event EventHandler? CatalogueChange;

        private readonly SemaphoreSlim _ecritureSema = new(1, 1);

        private static readonly JsonSerializerOptions _jsonOpts = new()
        {
            WriteIndented = true,
            PropertyNamingPolicy = null  // casse C# conservée
        };

        private CatalogueAppareilsService() { }

        public async Task ChargerAsync()
        {
            await _ecritureSema.WaitAsync();
            try
            {
                string fichier = CheminsMetrologo.FichierCatalogueAppareils;

                // Premier démarrage : reprise du JSON hérité s'il existe.
                if (!File.Exists(fichier))
                {
                    await ImporterJsonHeriteSiPresentAsync(fichier);
                }

                // Catalogue vide si le fichier n'existe toujours pas.
                List<ModeleAppareil> charges = await LireFichierAsync(fichier);

                Modeles.Clear();
                foreach (var m in charges) Modeles.Add(m);
                NotifierChange();
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Configuration, "CATALOGUE_LOAD_ERR",
                    $"Chargement catalogue échoué : {ex.Message}");
                Modeles.Clear();
                NotifierChange();
            }
            finally
            {
                _ecritureSema.Release();
            }
        }

        /// <summary>
        /// Ajoute un modèle en mémoire seulement (sans écriture réseau) s'il n'existe pas (même Id) :
        /// sert à semer les profils legacy sans polluer le catalogue partagé ni exiger M:\.
        /// Renvoie false s'il y était déjà.
        /// </summary>
        public bool AjouterEnMemoireSiAbsent(ModeleAppareil modele)
        {
            if (Modeles.Any(m => m.Id == modele.Id)) return false;
            Modeles.Add(modele);
            NotifierChange();
            return true;
        }

        /// <summary>
        /// Pose un profil en mémoire (sans toucher au réseau) : remplace l'entrée de même Id, sinon
        /// l'ajoute. Permet de resemer les profils legacy au démarrage en écrasant une copie
        /// dégradée d'appareils.json (ex. EIP sauvegardé dans l'éditeur générique, qui perd Legacy /
        /// AdresseFixeParDefaut / CommandesGateParSlot). Le profil de référence vient de
        /// appareils-legacy.json (cf. <see cref="SeedLegacyAppareils"/>) et a le dernier mot.
        /// </summary>
        public void RemplacerOuAjouterEnMemoire(ModeleAppareil modele)
        {
            for (int i = 0; i < Modeles.Count; i++)
            {
                if (Modeles[i].Id == modele.Id)
                {
                    Modeles[i] = modele;   // sur place : l'UI reçoit un événement Replace
                    NotifierChange();
                    return;
                }
            }
            Modeles.Add(modele);
            NotifierChange();
        }

        public async Task AjouterAsync(ModeleAppareil modele)
        {
            if (string.IsNullOrEmpty(modele.Id)) modele.Id = GenererId(modele.Nom);
            if (modele.DateCreation == default) modele.DateCreation = DateTime.Now;

            await _ecritureSema.WaitAsync();
            try
            {
                Modeles.Add(modele);
                await EcrireFichierAsync(CheminsMetrologo.FichierCatalogueAppareils, Modeles);
            }
            finally
            {
                _ecritureSema.Release();
            }
            NotifierChange();
        }

        /// <summary>
        /// Importe un ou plusieurs modèles depuis du JSON (objet seul ou tableau) : mise à jour si
        /// l'Id existe, ajout sinon. Le fichier n'est réécrit qu'une fois, en fin de lot.
        /// </summary>
        public async Task<int> ImporterDepuisJsonAsync(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                throw new ArgumentException("JSON vide.");

            var trimmed = json.TrimStart();
            List<ModeleAppareil>? modeles;
            if (trimmed.StartsWith("["))
            {
                modeles = JsonSerializer.Deserialize<List<ModeleAppareil>>(json, _jsonOpts);
            }
            else
            {
                var unique = JsonSerializer.Deserialize<ModeleAppareil>(json, _jsonOpts);
                modeles = unique != null ? new List<ModeleAppareil> { unique } : null;
            }

            if (modeles == null || modeles.Count == 0)
                throw new InvalidOperationException("JSON parsé en collection vide.");

            int n = 0;
            await _ecritureSema.WaitAsync();
            try
            {
                foreach (var m in modeles)
                {
                    if (string.IsNullOrEmpty(m.Id)) m.Id = GenererId(m.Nom);
                    if (string.IsNullOrEmpty(m.CreePar)) m.CreePar = "import";
                    if (m.DateCreation == default) m.DateCreation = DateTime.Now;

                    var existant = Modeles.FirstOrDefault(x => x.Id == m.Id);
                    if (existant != null)
                    {
                        existant.Nom = m.Nom;
                        existant.FabricantIdn = m.FabricantIdn;
                        existant.ModeleIdn = m.ModeleIdn;
                        existant.Parametres = m.Parametres;
                        existant.Gates = m.Gates;
                        existant.Entrees = m.Entrees;
                        existant.Couplages = m.Couplages;
                        existant.Reglages = m.Reglages;
                    }
                    else
                    {
                        Modeles.Add(m);
                    }
                    n++;
                }
                await EcrireFichierAsync(CheminsMetrologo.FichierCatalogueAppareils, Modeles);
            }
            finally
            {
                _ecritureSema.Release();
            }

            JournalLog.Info(CategorieLog.Administration, "CATALOGUE_IMPORT",
                $"{n} modèle(s) d'appareil importé(s) depuis JSON.");
            NotifierChange();
            return n;
        }

        public async Task ModifierAsync(string id, Action<ModeleAppareil> modification)
        {
            ModeleAppareil? cible;
            await _ecritureSema.WaitAsync();
            try
            {
                cible = Modeles.FirstOrDefault(m => m.Id == id);
                if (cible == null) return;
                modification(cible);
                await EcrireFichierAsync(CheminsMetrologo.FichierCatalogueAppareils, Modeles);
            }
            finally
            {
                _ecritureSema.Release();
            }
            NotifierChange();
        }

        public async Task SupprimerAsync(string id)
        {
            await _ecritureSema.WaitAsync();
            try
            {
                var cible = Modeles.FirstOrDefault(m => m.Id == id);
                if (cible == null) return;
                Modeles.Remove(cible);
                await EcrireFichierAsync(CheminsMetrologo.FichierCatalogueAppareils, Modeles);
            }
            finally
            {
                _ecritureSema.Release();
            }
            NotifierChange();
        }

        public ModeleAppareil? TrouverParIdn(string? fabricant, string? modele)
            => Modeles.FirstOrDefault(m => m.Correspond(fabricant, modele));

        /// <summary>
        /// Lit le JSON du catalogue. Liste vide si le fichier est absent ou corrompu (incident
        /// journalisé, sans exception : l'app reste utilisable).
        /// </summary>
        private static async Task<List<ModeleAppareil>> LireFichierAsync(string chemin)
        {
            if (!File.Exists(chemin)) return new List<ModeleAppareil>();

            try
            {
                string json = await File.ReadAllTextAsync(chemin);
                if (string.IsNullOrWhiteSpace(json)) return new List<ModeleAppareil>();
                var liste = JsonSerializer.Deserialize<List<ModeleAppareil>>(json, _jsonOpts);
                return liste ?? new List<ModeleAppareil>();
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Configuration, "CATALOGUE_JSON_CORROMPU",
                    $"Fichier catalogue corrompu ({Path.GetFileName(chemin)}) : {ex.Message}. "
                  + "L'app démarre sur un catalogue vide — il faudra restaurer une sauvegarde à la main.");
                return new List<ModeleAppareil>();
            }
        }

        /// <summary>
        /// Écriture atomique (.tmp puis <see cref="File.Replace"/>) : un crash en cours d'écriture
        /// ne laisse pas de fichier tronqué sur le partage.
        /// </summary>
        private static async Task EcrireFichierAsync(string chemin, IEnumerable<ModeleAppareil> modeles)
        {
            string? dossier = Path.GetDirectoryName(chemin);
            if (!string.IsNullOrEmpty(dossier)) Directory.CreateDirectory(dossier);

            string tmp = chemin + ".tmp";
            string json = JsonSerializer.Serialize(modeles.ToList(), _jsonOpts);
            await File.WriteAllTextAsync(tmp, json);

            if (File.Exists(chemin))
            {
                // Replace est atomique sur NTFS ; si M:\ est sur un autre système de fichiers,
                // repli sur Delete+Move (non atomique).
                try { File.Replace(tmp, chemin, destinationBackupFileName: null); }
                catch (PlatformNotSupportedException)
                {
                    File.Delete(chemin);
                    File.Move(tmp, chemin);
                }
            }
            else
            {
                File.Move(tmp, chemin);
            }
        }

        /// <summary>
        /// Migration : si le fichier réseau n'existe pas, reprend un JSON hérité (dossiers Catalogues
        /// local et réseau) : <c>AppareilsCatalogue.json</c>, ou à défaut le plus récent backup
        /// <c>AppareilsCatalogue.json.imported.*</c> (laissé par l'ancienne migration SQL, qui
        /// renommait l'original sans recréer le fichier actif). Les sources restent en place.
        /// </summary>
        private async Task ImporterJsonHeriteSiPresentAsync(string fichierCible)
        {
            string? cheminSource = TrouverFichierLegacy();
            if (cheminSource == null) return;

            try
            {
                string contenu = await File.ReadAllTextAsync(cheminSource);
                var liste = JsonSerializer.Deserialize<List<ModeleAppareil>>(contenu, _jsonOpts);
                if (liste == null || liste.Count == 0) return;

                foreach (var m in liste)
                {
                    if (string.IsNullOrEmpty(m.Id)) m.Id = GenererId(m.Nom);
                    if (m.DateCreation == default) m.DateCreation = DateTime.Now;
                }

                await EcrireFichierAsync(fichierCible, liste);

                JournalLog.Info(CategorieLog.Configuration, "CATALOGUE_MIGRE_VERS_RESEAU",
                    $"{liste.Count} modèle(s) récupéré(s) depuis {Path.GetFileName(cheminSource)} "
                  + $"et écrit(s) dans {fichierCible}.");
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Configuration, "CATALOGUE_MIGRATION_ERR",
                    $"Récupération depuis {Path.GetFileName(cheminSource)} échouée : {ex.Message}. "
                  + "Le fichier source reste là où il est, pour qu'on puisse l'analyser à la main.");
            }
        }

        /// <summary>
        /// Cherche un JSON hérité exploitable, dans l'ordre : <c>AppareilsCatalogue.json</c> (avec
        /// repli sur l'ancien chemin à plat de <c>%LocalAppData%\Metrologo\</c>), puis le plus
        /// récent <c>AppareilsCatalogue.json.imported.*</c> du dossier Catalogues, puis du dossier
        /// local. Renvoie null si rien ne convient.
        /// </summary>
        private static string? TrouverFichierLegacy()
        {
            string actif = CheminsMetrologo.ResoudreCheminAvecFallback(
                CheminsMetrologo.FichierAppareilsCatalogueLegacy, "AppareilsCatalogue.json");
            if (File.Exists(actif)) return actif;

            // Backups .imported.* du dossier Catalogues courant.
            string dossierCatalogues = CheminsMetrologo.Catalogues;
            if (Directory.Exists(dossierCatalogues))
            {
                var backups = Directory.EnumerateFiles(dossierCatalogues,
                                                       "AppareilsCatalogue.json.imported.*")
                                       .OrderByDescending(f => File.GetLastWriteTimeUtc(f))
                                       .ToList();
                if (backups.Count > 0) return backups[0];
            }

            // Idem dans le dossier local (Catalogues pointe sur le réseau, vieux backup resté en local).
            string dossierLocal = Path.Combine(CheminsMetrologo.Racine, "Catalogues");
            if (Directory.Exists(dossierLocal))
            {
                var backupsLocaux = Directory.EnumerateFiles(dossierLocal,
                                                              "AppareilsCatalogue.json.imported.*")
                                              .OrderByDescending(f => File.GetLastWriteTimeUtc(f))
                                              .ToList();
                if (backupsLocaux.Count > 0) return backupsLocaux[0];
            }

            return null;
        }

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
