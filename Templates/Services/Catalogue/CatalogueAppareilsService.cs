using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Dapper;
using Metrologo.Models;
using Metrologo.Services.Journal;
using Microsoft.Data.SqlClient;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Catalogue
{
    /// <summary>
    /// Catalogue centralisé des modèles d'appareils enregistrés (SCPI / IEEE).
    ///
    /// Stockage : table <c>dbo.T_CATALOGUE_APPAREILS</c> sur SQL Server (champs
    /// cherchables en colonnes, sous-collections complexes — Parametres, Gates,
    /// Entrees, Couplages, Reglages — sérialisées en JSON dans une colonne
    /// <c>Configuration</c>).
    ///
    /// Au premier démarrage avec une base vide, l'éventuel fichier JSON local
    /// hérité (<c>%LocalAppData%\Metrologo\AppareilsCatalogue.json</c>) est importé
    /// puis renommé en <c>.imported</c> pour ne pas être ré-importé. Migration
    /// transparente : l'utilisateur ne perd pas ses modèles existants en passant
    /// du mode JSON local au mode centralisé.
    ///
    /// Singleton accessible via <see cref="Instance"/>. <see cref="Modeles"/>
    /// est rechargée en mémoire à chaque <see cref="ChargerAsync"/> et bindable
    /// directement depuis l'UI ; <see cref="CatalogueChange"/> notifie l'UI à
    /// chaque ajout / modification / suppression / rechargement.
    /// </summary>
    public class CatalogueAppareilsService
    {
        private static readonly Lazy<CatalogueAppareilsService> _instance = new(() => new());
        public static CatalogueAppareilsService Instance => _instance.Value;

        public ObservableCollection<ModeleAppareil> Modeles { get; } = new();
        public event EventHandler? CatalogueChange;

        private static readonly JsonSerializerOptions _jsonOpts = new()
        {
            WriteIndented = false,
            PropertyNamingPolicy = null  // garde la casse C#, plus simple à debugger
        };

        private CatalogueAppareilsService() { }

        // -------------------------------------------------------------------------
        // Chargement
        // -------------------------------------------------------------------------

        public async Task ChargerAsync()
        {
            try
            {
                using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
                await c.OpenAsync();

                // Migration one-shot : si la table est vide ET qu'un JSON hérité existe,
                // on importe avant de charger. Idempotent : le JSON est renommé après.
                int nbModeles = await c.ExecuteScalarAsync<int>(
                    "SELECT COUNT(*) FROM dbo.T_CATALOGUE_APPAREILS");
                if (nbModeles == 0)
                {
                    await ImporterJsonHeriteSiPresentAsync(c);
                }

                var rows = await c.QueryAsync<RowAppareil>(
                    @"SELECT Id, Nom, FabricantIdn, ModeleIdn, Configuration, DateCreation, CreePar
                      FROM dbo.T_CATALOGUE_APPAREILS
                      ORDER BY Nom");

                Modeles.Clear();
                foreach (var r in rows)
                {
                    var m = HydraterDepuisRow(r);
                    if (m != null) Modeles.Add(m);
                }
                NotifierChange();
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Configuration, "CATALOGUE_LOAD_ERR",
                    $"Chargement catalogue échoué : {ex.Message}");
                Modeles.Clear();
                NotifierChange();
            }
        }

        // -------------------------------------------------------------------------
        // CRUD
        // -------------------------------------------------------------------------

        public async Task AjouterAsync(ModeleAppareil modele)
        {
            if (string.IsNullOrEmpty(modele.Id)) modele.Id = GenererId(modele.Nom);

            using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
            await c.OpenAsync();
            await c.ExecuteAsync(
                @"INSERT INTO dbo.T_CATALOGUE_APPAREILS
                  (Id, Nom, FabricantIdn, ModeleIdn, Configuration, DateCreation, CreePar)
                  VALUES (@Id, @Nom, @FabricantIdn, @ModeleIdn, @Configuration, @DateCreation, @CreePar)",
                new
                {
                    modele.Id,
                    modele.Nom,
                    modele.FabricantIdn,
                    modele.ModeleIdn,
                    Configuration = SerialiserConfiguration(modele),
                    DateCreation = DateTime.UtcNow,
                    modele.CreePar
                });

            Modeles.Add(modele);
            NotifierChange();
        }

        public async Task ModifierAsync(string id, Action<ModeleAppareil> modification)
        {
            var cible = Modeles.FirstOrDefault(m => m.Id == id);
            if (cible == null) return;

            modification(cible);

            using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
            await c.OpenAsync();
            await c.ExecuteAsync(
                @"UPDATE dbo.T_CATALOGUE_APPAREILS
                  SET Nom = @Nom,
                      FabricantIdn = @FabricantIdn,
                      ModeleIdn = @ModeleIdn,
                      Configuration = @Configuration,
                      DateModif = SYSUTCDATETIME()
                  WHERE Id = @Id",
                new
                {
                    cible.Id,
                    cible.Nom,
                    cible.FabricantIdn,
                    cible.ModeleIdn,
                    Configuration = SerialiserConfiguration(cible)
                });

            NotifierChange();
        }

        public async Task SupprimerAsync(string id)
        {
            var cible = Modeles.FirstOrDefault(m => m.Id == id);
            if (cible == null) return;

            using var c = new SqlConnection(MetrologoDbConnection.ConnectionString);
            await c.OpenAsync();
            await c.ExecuteAsync("DELETE FROM dbo.T_CATALOGUE_APPAREILS WHERE Id = @Id", new { Id = id });

            Modeles.Remove(cible);
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
        // Migration JSON hérité
        // -------------------------------------------------------------------------

        private async Task ImporterJsonHeriteSiPresentAsync(SqlConnection c)
        {
            string cheminJson = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Metrologo", "AppareilsCatalogue.json");
            if (!File.Exists(cheminJson)) return;

            try
            {
                string contenu = await File.ReadAllTextAsync(cheminJson);
                var liste = JsonSerializer.Deserialize<List<ModeleAppareil>>(contenu);
                if (liste == null || liste.Count == 0) return;

                int nbImportes = 0;
                foreach (var m in liste)
                {
                    if (string.IsNullOrEmpty(m.Id)) m.Id = GenererId(m.Nom);
                    try
                    {
                        await c.ExecuteAsync(
                            @"INSERT INTO dbo.T_CATALOGUE_APPAREILS
                              (Id, Nom, FabricantIdn, ModeleIdn, Configuration, DateCreation, CreePar)
                              VALUES (@Id, @Nom, @FabricantIdn, @ModeleIdn, @Configuration, @DateCreation, @CreePar)",
                            new
                            {
                                m.Id,
                                m.Nom,
                                m.FabricantIdn,
                                m.ModeleIdn,
                                Configuration = SerialiserConfiguration(m),
                                DateCreation = m.DateCreation == default ? DateTime.UtcNow : m.DateCreation,
                                m.CreePar
                            });
                        nbImportes++;
                    }
                    catch (Exception ex)
                    {
                        JournalLog.Warn(CategorieLog.Configuration, "CATALOGUE_IMPORT_SKIP",
                            $"Import du modèle {m.Nom} échoué : {ex.Message}");
                    }
                }

                // Renomme le JSON pour qu'il ne soit pas ré-importé au prochain lancement.
                string cible = cheminJson + ".imported." + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss");
                File.Move(cheminJson, cible);

                JournalLog.Info(CategorieLog.Configuration, "CATALOGUE_MIGRE_VERS_SQL",
                    $"{nbImportes} modèle(s) importé(s) depuis le JSON local. Backup : {Path.GetFileName(cible)}.");
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Configuration, "CATALOGUE_IMPORT_ERR",
                    $"Migration JSON → SQL échouée : {ex.Message}. Le JSON reste en place pour analyse manuelle.");
            }
        }

        // -------------------------------------------------------------------------
        // (Dé)sérialisation Configuration
        // -------------------------------------------------------------------------

        /// <summary>
        /// Sérialise la partie « configuration » d'un modèle (tout sauf les colonnes
        /// plates Id/Nom/IDN/dates) en JSON. Format compact, pas indenté — pour
        /// stockage SQL.
        /// </summary>
        private static string SerialiserConfiguration(ModeleAppareil m)
        {
            var payload = new ConfigurationPayload
            {
                Parametres = m.Parametres,
                Gates = m.Gates,
                Entrees = m.Entrees,
                Couplages = m.Couplages,
                Reglages = m.Reglages
            };
            return JsonSerializer.Serialize(payload, _jsonOpts);
        }

        private static ModeleAppareil? HydraterDepuisRow(RowAppareil r)
        {
            try
            {
                var payload = JsonSerializer.Deserialize<ConfigurationPayload>(
                    r.Configuration ?? "{}", _jsonOpts) ?? new ConfigurationPayload();

                return new ModeleAppareil
                {
                    Id = r.Id,
                    Nom = r.Nom,
                    FabricantIdn = r.FabricantIdn ?? "",
                    ModeleIdn = r.ModeleIdn ?? "",
                    DateCreation = DateTime.SpecifyKind(r.DateCreation, DateTimeKind.Utc).ToLocalTime(),
                    CreePar = r.CreePar ?? "",
                    Parametres = payload.Parametres ?? new ParametresIeee(),
                    Gates = payload.Gates ?? new List<string>(),
                    Entrees = payload.Entrees ?? new List<string>(),
                    Couplages = payload.Couplages ?? new List<string>(),
                    Reglages = payload.Reglages ?? new List<ReglageAppareil>()
                };
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Configuration, "CATALOGUE_HYDRATE_ERR",
                    $"Modèle {r.Id} : Configuration JSON corrompue ({ex.Message}) — modèle ignoré.");
                return null;
            }
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

        // DTO interne Dapper
        private sealed class RowAppareil
        {
            public string Id { get; set; } = "";
            public string Nom { get; set; } = "";
            public string? FabricantIdn { get; set; }
            public string? ModeleIdn { get; set; }
            public string? Configuration { get; set; }
            public DateTime DateCreation { get; set; }
            public string? CreePar { get; set; }
        }

        // Sous-payload sérialisé dans la colonne Configuration
        private sealed class ConfigurationPayload
        {
            public ParametresIeee? Parametres { get; set; }
            public List<string>? Gates { get; set; }
            public List<string>? Entrees { get; set; }
            public List<string>? Couplages { get; set; }
            public List<ReglageAppareil>? Reglages { get; set; }
        }
    }
}
