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
    /// Le catalogue central des modèles d'appareils SCPI/IEEE. Tout tient dans un unique fichier
    /// JSON posé sur le partage réseau (par défaut M:\exe_spe\Data_Metrologo\Catalogues\appareils.json) :
    /// on le relit à chaque ChargerAsync et on le réécrit de façon atomique (temp + Move) après
    /// chaque modif. La toute première fois, si le fichier réseau n'est pas là, on récupère l'ancien
    /// JSON local (%LocalAppData%\Metrologo\Catalogues\AppareilsCatalogue.json) hérité de l'époque
    /// SQL Server. Côté multi-postes : un SemaphoreSlim plus l'écriture atomique nous mettent à l'abri
    /// de la corruption, mais c'est le dernier qui écrit qui l'emporte. Pour gérer un vrai merge il
    /// faudrait passer par un fichier .lock dédié.
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
            PropertyNamingPolicy = null  // on garde la casse C# telle quelle, c'est plus simple à débugger
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

                // Tout premier démarrage : on récupère le JSON hérité local s'il traîne quelque part
                // (sinon ça ne fait rien). Du 100% fichier, plus aucune dépendance SQL.
                if (!File.Exists(fichier))
                {
                    await ImporterJsonHeriteSiPresentAsync(fichier);
                }

                // Le fichier réseau (qui peut très bien ne pas exister si la migration n'a rien
                // déniché : dans ce cas on démarre sur un catalogue vide)
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
        /// Ajoute un modèle en mémoire seulement (pas d'écriture réseau), à condition qu'il ne soit
        /// pas déjà là (même Id). Ça nous sert à semer les profils legacy au démarrage sans polluer
        /// le catalogue partagé ni avoir besoin de M:\. Renvoie false s'il y était déjà.
        /// </summary>
        public bool AjouterEnMemoireSiAbsent(ModeleAppareil modele)
        {
            if (Modeles.Any(m => m.Id == modele.Id)) return false;
            Modeles.Add(modele);
            NotifierChange();
            return true;
        }

        /// <summary>
        /// Pose un profil en mémoire (sans toucher au réseau) : si une entrée de même Id existe déjà
        /// on la remplace, sinon on l'ajoute. C'est ce qui nous permet de (re)semer les profils legacy
        /// au démarrage en ÉCRASANT une éventuelle copie corrompue qui se serait glissée dans
        /// appareils.json — typiquement un EIP qu'on aurait ouvert puis sauvegardé par mégarde dans
        /// l'éditeur générique, et qui aurait perdu au passage Legacy / AdresseFixeParDefaut /
        /// CommandesGateParSlot. Le profil de référence, lui, vient de appareils-legacy.json
        /// (cf. <see cref="SeedLegacyAppareils"/>), donc on lui donne systématiquement le dernier mot
        /// sur ce qui pourrait traîner dans le catalogue principal.
        /// </summary>
        public void RemplacerOuAjouterEnMemoire(ModeleAppareil modele)
        {
            for (int i = 0; i < Modeles.Count; i++)
            {
                if (Modeles[i].Id == modele.Id)
                {
                    Modeles[i] = modele;   // on remplace sur place (l'UI reçoit un event Replace)
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
        /// Importe un ou plusieurs modèles à partir de JSON (un objet seul ou un tableau). Si l'Id
        /// existe déjà on met à jour, sinon on ajoute. Le fichier n'est réécrit qu'une fois, à la
        /// fin du batch.
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
        /// Lit le fichier JSON et désérialise la liste de modèles. On renvoie une liste vide si le
        /// fichier est absent ou corrompu : dans ce dernier cas on log l'incident mais on ne lève
        /// rien, histoire que l'app reste utilisable.
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
        /// Écriture atomique : on écrit d'abord dans un fichier .tmp, puis on bascule sur l'original
        /// via <see cref="File.Replace"/>. Comme ça, un crash en plein milieu de l'écriture ne peut
        /// pas laisser un fichier à moitié écrit (donc corrompu) sur le partage réseau.
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
                // Replace est atomique sur NTFS. Si jamais le partage M:\ tourne sur autre chose,
                // on se rabat sur Delete+Move : nettement moins safe, mais ça dépanne.
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
        /// Au tout premier démarrage avec le nouveau stockage fichier, si le fichier réseau n'est
        /// pas là, on part à la pêche au JSON hérité dans les dossiers Catalogues (local + réseau),
        /// sous l'une ou l'autre de ces formes :
        ///   • <c>AppareilsCatalogue.json</c>           — le format actif
        ///   • <c>AppareilsCatalogue.json.imported.*</c> — un backup laissé par une migration
        ///     antérieure (cas classique : la migration SQL d'avril qui a renommé l'original en
        ///     .imported.YYYYMMDD-HHMMSS sans recréer le fichier actif)
        /// On prend le plus récent qui contient des données valables, on l'écrit dans le fichier
        /// réseau cible, et on laisse les backups là où ils sont au cas où on doive les inspecter.
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

                // On comble les Id manquants au passage
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
        /// Part en quête d'un JSON hérité exploitable, en suivant cet ordre :
        ///   1. <c>AppareilsCatalogue.json</c> (le format actif) dans le dossier Catalogues
        ///   2. à défaut, le plus récent <c>AppareilsCatalogue.json.imported.*</c> du dossier
        ///      Catalogues (un backup d'une migration passée — typiquement avril 2026)
        ///   3. en dernier recours, l'ancien chemin à plat dans <c>%LocalAppData%\Metrologo\</c>
        /// Renvoie null si rien ne convient.
        /// </summary>
        private static string? TrouverFichierLegacy()
        {
            // 1. Le format actif (sans le suffixe .imported)
            string actif = CheminsMetrologo.ResoudreCheminAvecFallback(
                CheminsMetrologo.FichierAppareilsCatalogueLegacy, "AppareilsCatalogue.json");
            if (File.Exists(actif)) return actif;

            // 2. Les backups .imported.* (dans le dossier Catalogues courant = M:\ par défaut)
            string dossierCatalogues = CheminsMetrologo.Catalogues;
            if (Directory.Exists(dossierCatalogues))
            {
                var backups = Directory.EnumerateFiles(dossierCatalogues,
                                                       "AppareilsCatalogue.json.imported.*")
                                       .OrderByDescending(f => File.GetLastWriteTimeUtc(f))
                                       .ToList();
                if (backups.Count > 0) return backups[0];
            }

            // 3. Pareil, mais dans le dossier local (le cas où Catalogues pointe sur le réseau
            // alors qu'un vieux backup est resté sur la machine locale)
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
