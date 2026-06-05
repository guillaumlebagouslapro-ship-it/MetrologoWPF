using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Metrologo.Models;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Catalogue
{
    /// <summary>
    /// Catalogue centralisé des modèles d'appareils enregistrés (SCPI / IEEE).
    ///
    /// Stockage : fichier JSON unique sur le partage réseau —
    /// <see cref="CheminsMetrologo.FichierCatalogueAppareils"/>
    /// (par défaut <c>M:\exe_spe\Data_Metrologo\Catalogues\appareils.json</c>). Lecture
    /// au démarrage de l'app + à chaque <see cref="ChargerAsync"/> ; ré-écriture
    /// atomique (fichier temp + Move) après chaque modification, partage entre postes
    /// transparent.
    ///
    /// Migration depuis l'ancien stockage SQL Server : si le fichier réseau n'existe
    /// pas encore mais qu'un JSON local hérité est présent
    /// (<c>%LocalAppData%\Metrologo\Catalogues\AppareilsCatalogue.json</c>), il est
    /// importé automatiquement au premier chargement.
    ///
    /// Concurrence multi-postes : verrou applicatif via SemaphoreSlim + écriture
    /// atomique. Deux postes qui modifient simultanément ne corrompront pas le
    /// fichier (le dernier qui écrit gagne). Pour des scénarios plus complexes
    /// (merge), il faudra introduire un fichier .lock dédié.
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
            PropertyNamingPolicy = null  // garde la casse C#, plus simple à debugger
        };

        private CatalogueAppareilsService() { }

        // -------------------------------------------------------------------------
        // Chargement
        // -------------------------------------------------------------------------

        public async Task ChargerAsync()
        {
            await _ecritureSema.WaitAsync();
            try
            {
                string fichier = CheminsMetrologo.FichierCatalogueAppareils;

                // 1ʳᵉ exécution : tenter une migration depuis le JSON hérité local
                // (idempotente : no-op si le fichier réseau existe déjà ou si pas de
                // legacy local trouvé). 100 % fichier — pas de dépendance SQL.
                if (!File.Exists(fichier))
                {
                    await ImporterJsonHeriteSiPresentAsync(fichier);
                }

                // Lecture du fichier réseau (peut rester absent si la migration n'a
                // rien trouvé ou si le fichier vient d'être supprimé — dans ce cas
                // on démarre sur un catalogue vide).
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

        // -------------------------------------------------------------------------
        // CRUD
        // -------------------------------------------------------------------------

        /// <summary>
        /// Ajoute un modèle <b>en mémoire uniquement</b> (aucune écriture sur le partage réseau)
        /// s'il n'existe pas déjà (comparaison par <c>Id</c>). Utilisé pour seeder les profils
        /// legacy au démarrage : ils sont définis dans le code et toujours présents, sans polluer
        /// le catalogue réseau partagé ni dépendre de la disponibilité de <c>M:\</c>.
        /// Retourne <c>true</c> si le modèle a été ajouté, <c>false</c> s'il était déjà présent.
        /// </summary>
        public bool AjouterEnMemoireSiAbsent(ModeleAppareil modele)
        {
            if (Modeles.Any(m => m.Id == modele.Id)) return false;
            Modeles.Add(modele);
            NotifierChange();
            return true;
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
        /// Importe un ou plusieurs modèles depuis une chaîne JSON (objet seul ou tableau).
        /// Les Id existants sont mis à jour, les nouveaux sont ajoutés. Une seule
        /// ré-écriture du fichier à la fin de la batch.
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

        // -------------------------------------------------------------------------
        // Recherche IDN
        // -------------------------------------------------------------------------

        public ModeleAppareil? TrouverParIdn(string? fabricant, string? modele)
            => Modeles.FirstOrDefault(m => m.Correspond(fabricant, modele));

        public bool EstDansCatalogue(string? fabricant, string? modele)
            => TrouverParIdn(fabricant, modele) != null;

        // -------------------------------------------------------------------------
        // I/O fichier
        // -------------------------------------------------------------------------

        /// <summary>
        /// Lit le fichier JSON et désérialise la liste de modèles. Retourne une
        /// liste vide si le fichier n'existe pas ou est corrompu (la corruption
        /// est loggée mais ne lève pas — l'app reste utilisable).
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
                  + "L'app démarre sur un catalogue vide — restaure une sauvegarde manuellement.");
                return new List<ModeleAppareil>();
            }
        }

        /// <summary>
        /// Écriture atomique : on écrit d'abord dans un fichier .tmp puis on remplace
        /// l'original par <see cref="File.Replace"/>. Garantit qu'un crash en cours
        /// d'écriture ne laisse pas un fichier corrompu sur le partage réseau.
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
                // Replace est atomique sur NTFS — si le partage M:\ est sur autre chose,
                // on retombe sur Delete+Move qui est nettement moins sûr mais marche.
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

        // -------------------------------------------------------------------------
        // Migration JSON hérité local → fichier réseau
        // -------------------------------------------------------------------------

        /// <summary>
        /// Au tout premier démarrage avec le nouveau stockage fichier, si le fichier
        /// réseau n'existe pas, on cherche un JSON hérité dans les dossiers Catalogues
        /// (local + réseau) sous plusieurs formes possibles :
        ///   • <c>AppareilsCatalogue.json</c>           — format actif
        ///   • <c>AppareilsCatalogue.json.imported.*</c> — backup d'une migration
        ///     précédente (typique : la migration SQL d'avril a renommé l'original
        ///     en .imported.YYYYMMDD-HHMMSS sans préserver l'actif)
        /// On charge le plus récent qui contient des données valides, l'écrit dans le
        /// fichier réseau cible, puis on laisse les backups en place pour analyse.
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

                // Normalise les Id manquants
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
                  + "Le fichier source reste en place pour analyse manuelle.");
            }
        }

        /// <summary>
        /// Cherche un JSON hérité utilisable, dans cet ordre :
        ///   1. <c>AppareilsCatalogue.json</c> (format actif) dans le dossier Catalogues
        ///   2. Le plus récent <c>AppareilsCatalogue.json.imported.*</c> dans le dossier
        ///      Catalogues (backup d'une migration précédente — typique : avril 2026)
        ///   3. Fallback ancien chemin à plat dans <c>%LocalAppData%\Metrologo\</c>
        /// Retourne null si aucun candidat trouvé.
        /// </summary>
        private static string? TrouverFichierLegacy()
        {
            // 1. Format actif (sans suffixe .imported)
            string actif = CheminsMetrologo.ResoudreCheminAvecFallback(
                CheminsMetrologo.FichierAppareilsCatalogueLegacy, "AppareilsCatalogue.json");
            if (File.Exists(actif)) return actif;

            // 2. Backups .imported.* (dans le dossier Catalogues actuel = M:\ par défaut)
            string dossierCatalogues = CheminsMetrologo.Catalogues;
            if (Directory.Exists(dossierCatalogues))
            {
                var backups = Directory.EnumerateFiles(dossierCatalogues,
                                                       "AppareilsCatalogue.json.imported.*")
                                       .OrderByDescending(f => File.GetLastWriteTimeUtc(f))
                                       .ToList();
                if (backups.Count > 0) return backups[0];
            }

            // 3. Idem dans le dossier local (cas où Catalogues pointe sur le réseau mais
            // qu'il reste un backup historique sur la machine locale)
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

        // -------------------------------------------------------------------------
        // Helpers
        // -------------------------------------------------------------------------

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
