using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Metrologo.Services;

namespace Metrologo.Models
{
    /// <summary>
    /// Préférences de l'application, sur deux couches.
    /// LOCAL : %LocalAppData%\Metrologo\Configuration\settings.json (rubidium actif courant,
    /// chemin de la macro Excel perso). RESEAU : fichiers dédiés sur le partage M:\ pour ce qui
    /// est commun aux postes (Utilisateurs\utilisateurs.json avec comptes + hash mdp,
    /// Rubidiums\catalogue.json). Les fichiers réseau sont relus à la demande, avec cache
    /// mémoire invalidé au write, pour voir les modifs des autres postes sans redémarrer.
    /// </summary>
    public static class Preferences
    {
        private static Settings _settings = new();

        // -------- local : rubidium actif + macro --------

        /// <summary>
        /// Rubidium actif. Lu en priorité depuis le fichier partagé pour que tous les postes
        /// reprennent le même rubidium de référence ; repli sur le réglage local si le partage est injoignable.
        /// </summary>
        public static Rubidium? RubidiumActif =>
            LireFichierReseau<Rubidium>(CheminsMetrologo.FichierRubidiumActif) ?? _settings.RubidiumActif;

        /// <summary>Chemin par défaut du Metrologo.xla (nouvel emplacement FCT_VBA2016).</summary>
        private const string CheminMacroParDefaut = @"C:\EXE_SPE\FCT_VBA2016\Metrologo.xla";

        /// <summary>Ancien chemin par défaut, remplacé automatiquement par le nouveau.</summary>
        private const string CheminMacroLegacy = @"C:\Exe_Spe\Fct_VBA\Metrologo.xla";

        public static string CheminMacroXLA
        {
            get
            {
                var c = _settings.CheminMacroXLA;
                // Aucun réglage, ou ancien chemin par défaut → nouveau chemin par défaut.
                // Un chemin personnalisé (autre) est conservé tel quel.
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

            // Migration one-shot : si le settings.json local contenait encore des
            // données qui ont été déplacées vers le réseau (utilisateurs, catalogue
            // rubidiums), on les migre vers les fichiers réseau puis on les retire
            // du settings local pour éviter le drift entre les deux sources.
            MigrerDonneesPartageesVersReseau();
        }

        public static void SauvegarderRubidium(Rubidium? rubi)
        {
            // PARTAGÉ : tous les postes reprennent ce rubidium de référence au prochain
            // démarrage (lecture dans le getter ci-dessus). Sans ça, le changement restait
            // local au poste qui l'a fait.
            if (rubi != null)
                EcrireFichierReseau(CheminsMetrologo.FichierRubidiumActif, rubi);

            // LOCAL : repli si le partage est injoignable (poste hors réseau).
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
                // silencieux : la perte d'une préférence locale n'est pas bloquante
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

        /// <summary>Invalide le cache : la prochaine lecture de Utilisateurs relira le fichier JSON
        /// (comptes modifiés depuis un autre poste, sans redémarrer).</summary>
        public static void InvaliderCacheUtilisateurs() => _cacheUtilisateurs = null;

        // -------- RÉSEAU : catalogue rubidiums --------

        private static List<Rubidium>? _cacheCatalogueRubidiums;

        /// <summary>
        /// Catalogue des rubidiums prédéfinis partagé entre postes. Si le fichier
        /// réseau est vide / absent, un seed par défaut (Syref + Redondances) est
        /// renvoyé pour amorcer l'UI.
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

        /// <summary>Invalide le cache : la prochaine lecture de CatalogueRubidiums relira le fichier réseau.</summary>
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

        /// <summary>Écriture atomique (.tmp + Replace) pour éviter la corruption.</summary>
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
                // Si l'écriture réseau échoue (partage HS), on perd la modification.
                // Le cache mémoire garde l'état pour la session en cours.
            }
        }

        // -------- Migration vers le réseau --------

        /// <summary>One-shot : migre vers le réseau les données encore présentes dans le settings.json local
        /// (utilisateurs, catalogue rubidiums) puis vide les champs locaux. Idempotent.</summary>
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

        // -------- DTO de persistance --------

        private class Settings
        {
            public Rubidium? RubidiumActif { get; set; }
            public string? CheminMacroXLA { get; set; }

            // Champs legacy gardés pour la migration one-shot vers le réseau.
            // Une fois migrés ils valent null et ne sont plus sérialisés.
            public List<Rubidium>? CatalogueRubidiums { get; set; }
            public List<Utilisateur>? Utilisateurs { get; set; }
        }
    }
}
