using System.IO;
using System.Text.Json;
using Metrologo.Services;

namespace Metrologo.Models
{
    /// <summary>
    /// Préférences de l'application, sur deux couches.
    /// LOCAL : %LocalAppData%\Metrologo\Configuration\settings.json (rubidium actif, chemin de la
    /// macro Excel perso). RÉSEAU : fichiers sur le partage M:\ communs aux postes
    /// (Utilisateurs\utilisateurs.json avec comptes + hash mdp, Rubidiums\catalogue.json).
    /// Les fichiers réseau sont relus à la demande, avec un cache mémoire invalidé à chaque
    /// écriture : les modifs faites depuis un autre poste sont vues sans redémarrer.
    /// </summary>
    public static class Preferences
    {
        private static Settings _settings = new();

        // -------- local : rubidium actif + macro --------

        /// <summary>
        /// Le rubidium actif, lu d'abord dans le fichier partagé (même référence pour tous les
        /// postes) ; repli sur le réglage local si le partage est injoignable.
        /// </summary>
        public static Rubidium? RubidiumActif =>
            LireFichierReseau<Rubidium>(CheminsMetrologo.FichierRubidiumActif) ?? _settings.RubidiumActif;

        /// <summary>Chemin par défaut du Metrologo.xla (FCT_VBA2016).</summary>
        private const string CheminMacroParDefaut = @"C:\EXE_SPE\FCT_VBA2016\Metrologo.xla";

        /// <summary>Ancien chemin par défaut, remplacé automatiquement par le nouveau.</summary>
        private const string CheminMacroLegacy = @"C:\Exe_Spe\Fct_VBA\Metrologo.xla";

        public static string CheminMacroXLA
        {
            get
            {
                var c = _settings.CheminMacroXLA;
                // Rien de réglé ou ancien chemin par défaut : nouveau défaut. Chemin perso : conservé.
                if (string.IsNullOrWhiteSpace(c)
                    || string.Equals(c, CheminMacroLegacy, System.StringComparison.OrdinalIgnoreCase))
                    return CheminMacroParDefaut;
                return c;
            }
            set { _settings.CheminMacroXLA = value; SauvegarderLocal(); }
        }

        public static void Charger()
        {
            try
            {
                string chemin = CheminsMetrologo.FichierSettings;
                if (!File.Exists(chemin)) return;
                var json = File.ReadAllText(chemin);
                var parsed = JsonSerializer.Deserialize<Settings>(json);
                if (parsed != null) _settings = parsed;
            }
            catch
            {
                _settings = new Settings();
            }

            // Migration one-shot des données désormais stockées sur le réseau (utilisateurs,
            // catalogue rubidiums) : évite deux sources qui divergent.
            MigrerDonneesPartageesVersReseau();

            // Format des rapports : figé pour la session (démarrage ou actualisation de la config).
            RechargerFormatRapports();
        }

        public static void SauvegarderRubidium(Rubidium? rubi)
        {
            // Partagé : tous les postes reprendront ce rubidium à leur prochain démarrage.
            if (rubi != null)
                EcrireFichierReseau(CheminsMetrologo.FichierRubidiumActif, rubi);

            // Local : repli quand le partage est injoignable.
            _settings.RubidiumActif = rubi;
            SauvegarderLocal();
        }

        private static void SauvegarderLocal()
        {
            try
            {
                string chemin = CheminsMetrologo.FichierSettings;
                var dir = Path.GetDirectoryName(chemin);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
                var json = JsonSerializer.Serialize(_settings,
                    new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(chemin, json);
            }
            catch
            {
                // Silencieux : perdre une préférence locale n'est pas bloquant.
            }
        }

        // -------- RÉSEAU : utilisateurs --------

        private static List<Utilisateur>? _cacheUtilisateurs;

        public static List<Utilisateur> Utilisateurs
        {
            get
            {
                _cacheUtilisateurs ??= LireFichierReseau<List<Utilisateur>>(
                    CheminsMetrologo.FichierUtilisateurs) ?? new List<Utilisateur>();
                return _cacheUtilisateurs;
            }
        }

        public static void SauvegarderUtilisateurs(IEnumerable<Utilisateur> liste)
        {
            var nouvelleListe = new List<Utilisateur>(liste);
            EcrireFichierReseau(CheminsMetrologo.FichierUtilisateurs, nouvelleListe);
            _cacheUtilisateurs = nouvelleListe;
        }

        /// <summary>Vide le cache : la prochaine lecture relit le fichier (comptes modifiés depuis un autre poste).</summary>
        public static void InvaliderCacheUtilisateurs() => _cacheUtilisateurs = null;

        // -------- RÉSEAU : catalogue rubidiums --------

        private static List<Rubidium>? _cacheCatalogueRubidiums;

        /// <summary>
        /// Catalogue des rubidiums, partagé entre les postes. Fichier réseau vide ou absent :
        /// renvoie un jeu par défaut (Syref + Redondances).
        /// </summary>
        public static List<Rubidium> CatalogueRubidiums
        {
            get
            {
                _cacheCatalogueRubidiums ??= LireFichierReseau<List<Rubidium>>(
                    CheminsMetrologo.FichierCatalogueRubidiums);

                if (_cacheCatalogueRubidiums == null || _cacheCatalogueRubidiums.Count == 0)
                {
                    return new List<Rubidium>
                    {
                        new Rubidium { Id = 1, Designation = "Syref",       FrequenceMoyenne = 10_000_000.0 },
                        new Rubidium { Id = 2, Designation = "Redondances", FrequenceMoyenne = 10_000_000.0 },
                    };
                }
                return _cacheCatalogueRubidiums;
            }
        }

        public static void SauvegarderCatalogueRubidiums(IEnumerable<Rubidium> catalogue)
        {
            var nouvelleListe = new List<Rubidium>(catalogue);
            EcrireFichierReseau(CheminsMetrologo.FichierCatalogueRubidiums, nouvelleListe);
            _cacheCatalogueRubidiums = nouvelleListe;
        }

        /// <summary>Vide le cache : la prochaine lecture relit le fichier réseau.</summary>
        public static void InvaliderCacheCatalogueRubidiums() => _cacheCatalogueRubidiums = null;

        // -------- I/O réseau --------

        private static T? LireFichierReseau<T>(string chemin) where T : class
        {
            try
            {
                if (!File.Exists(chemin)) return null;
                string json = File.ReadAllText(chemin);
                if (string.IsNullOrWhiteSpace(json)) return null;
                return JsonSerializer.Deserialize<T>(json);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>Écriture atomique (.tmp puis Replace) pour ne pas corrompre le fichier.</summary>
        private static void EcrireFichierReseau<T>(string chemin, T donnees)
        {
            try
            {
                string? dossier = Path.GetDirectoryName(chemin);
                if (!string.IsNullOrEmpty(dossier)) Directory.CreateDirectory(dossier);

                string tmp = chemin + ".tmp";
                string json = JsonSerializer.Serialize(donnees,
                    new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(tmp, json);

                if (File.Exists(chemin))
                {
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
            catch
            {
                // Partage HS : la modification est perdue côté fichier ; le cache mémoire
                // garde l'état le temps de la session.
            }
        }

        // -------- Migration vers le réseau --------

        /// <summary>Pousse vers le réseau les données restées dans le settings.json local
        /// (utilisateurs, catalogue rubidiums), puis vide les champs locaux. Idempotent.</summary>
        private static void MigrerDonneesPartageesVersReseau()
        {
            bool localModifie = false;

            // Utilisateurs
            if (_settings.Utilisateurs != null && _settings.Utilisateurs.Count > 0)
            {
                if (!File.Exists(CheminsMetrologo.FichierUtilisateurs))
                {
                    EcrireFichierReseau(CheminsMetrologo.FichierUtilisateurs, _settings.Utilisateurs);
                    _cacheUtilisateurs = _settings.Utilisateurs;
                }
                _settings.Utilisateurs = null;
                localModifie = true;
            }

            // Catalogue rubidiums
            if (_settings.CatalogueRubidiums != null && _settings.CatalogueRubidiums.Count > 0)
            {
                if (!File.Exists(CheminsMetrologo.FichierCatalogueRubidiums))
                {
                    EcrireFichierReseau(CheminsMetrologo.FichierCatalogueRubidiums, _settings.CatalogueRubidiums);
                    _cacheCatalogueRubidiums = _settings.CatalogueRubidiums;
                }
                _settings.CatalogueRubidiums = null;
                localModifie = true;
            }

            if (localModifie) SauvegarderLocal();
        }

        // -------- RÉSEAU : format des rapports Excel --------

        private static FormatRapport? _formatRapportsSession;

        /// <summary>
        /// Format appliqué par CE poste : lu au démarrage (<see cref="Charger"/>) et figé pour la
        /// session — un changement fait dans l'Admin d'un autre poste n'est pris qu'au prochain
        /// démarrage (ou via « charger la dernière configuration », cf. <see cref="RechargerFormatRapports"/>).
        /// </summary>
        public static FormatRapport FormatRapportsActif => _formatRapportsSession ??= FormatRapportsPartage;

        /// <summary>Relit le réglage partagé et l'applique à la session en cours.</summary>
        public static void RechargerFormatRapports() => _formatRapportsSession = FormatRapportsPartage;

        /// <summary>
        /// Réglage enregistré pour tous les postes : format des rapports Excel en sortie de mesure
        /// (hors stabilité). <c>.xls</c> par défaut : le logiciel externe d'extraction ne lit encore
        /// que ce format. Lu dans le fichier partagé, avec repli sur la dernière valeur connue du
        /// poste si le partage est injoignable.
        /// </summary>
        public static FormatRapport FormatRapportsPartage
        {
            get => LireFichierReseau<ReglagesRapports>(CheminsMetrologo.FichierReglagesRapports)?.Format
                   ?? _settings.FormatRapports
                   ?? FormatRapport.Xls;
            set
            {
                EcrireFichierReseau(CheminsMetrologo.FichierReglagesRapports, new ReglagesRapports { Format = value });
                _settings.FormatRapports = value;
                SauvegarderLocal();
            }
        }

        private class ReglagesRapports
        {
            [System.Text.Json.Serialization.JsonConverter(typeof(System.Text.Json.Serialization.JsonStringEnumConverter))]
            public FormatRapport Format { get; set; } = FormatRapport.Xls;
        }

        private class Settings
        {
            public Rubidium? RubidiumActif { get; set; }
            public string? CheminMacroXLA { get; set; }
            public FormatRapport? FormatRapports { get; set; }

            // Champs legacy conservés pour la migration vers le réseau (mis à null une fois migrés).
            public List<Rubidium>? CatalogueRubidiums { get; set; }
            public List<Utilisateur>? Utilisateurs { get; set; }
        }
    }
}
