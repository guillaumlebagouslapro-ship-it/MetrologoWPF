using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Metrologo.Services;

namespace Metrologo.Models
{
    /// <summary>
    /// Préférences de l'application, réparties sur deux couches.
    /// En LOCAL : %LocalAppData%\Metrologo\Configuration\settings.json (le rubidium actif du moment,
    /// le chemin de la macro Excel perso). Sur le RESEAU : des fichiers dédiés sur le partage M:\ pour
    /// tout ce qui est commun aux postes (Utilisateurs\utilisateurs.json avec les comptes + hash mdp,
    /// Rubidiums\catalogue.json). On relit les fichiers réseau à la demande, avec un cache mémoire
    /// qu'on invalide à chaque write, comme ça on voit les modifs faites depuis un autre poste sans
    /// avoir à redémarrer.
    /// </summary>
    public static class Preferences
    {
        private static Settings _settings = new();

        // -------- local : rubidium actif + macro --------

        /// <summary>
        /// Le rubidium actif. On le lit d'abord dans le fichier partagé pour que tous les postes
        /// repartent du même rubidium de référence ; si le partage est injoignable, on retombe sur le réglage local.
        /// </summary>
        public static Rubidium? RubidiumActif =>
            LireFichierReseau<Rubidium>(CheminsMetrologo.FichierRubidiumActif) ?? _settings.RubidiumActif;

        /// <summary>Chemin par défaut du Metrologo.xla, dans le nouvel emplacement FCT_VBA2016.</summary>
        private const string CheminMacroParDefaut = @"C:\EXE_SPE\FCT_VBA2016\Metrologo.xla";

        /// <summary>L'ancien chemin par défaut, qu'on remplace automatiquement par le nouveau.</summary>
        private const string CheminMacroLegacy = @"C:\Exe_Spe\Fct_VBA\Metrologo.xla";

        public static string CheminMacroXLA
        {
            get
            {
                var c = _settings.CheminMacroXLA;
                // Rien de réglé, ou bien l'ancien chemin par défaut : on bascule sur le nouveau.
                // Si l'utilisateur a mis un chemin perso (autre chose), on le garde tel quel.
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

            // Migration one-shot : si le settings.json local traîne encore des données
            // qui ont depuis été déplacées vers le réseau (utilisateurs, catalogue
            // rubidiums), on les pousse vers les fichiers réseau puis on les enlève du
            // settings local, pour ne pas se retrouver avec deux sources qui divergent.
            MigrerDonneesPartageesVersReseau();
        }

        public static void SauvegarderRubidium(Rubidium? rubi)
        {
            // CÔTÉ PARTAGÉ : tous les postes reprendront ce rubidium de référence à leur
            // prochain démarrage (c'est le getter ci-dessus qui le relit). Sans ça, le
            // changement serait resté cantonné au poste qui l'a fait.
            if (rubi != null)
                EcrireFichierReseau(CheminsMetrologo.FichierRubidiumActif, rubi);

            // CÔTÉ LOCAL : sert de repli quand le partage est injoignable (poste hors réseau).
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
                // on reste silencieux : perdre une préférence locale n'a rien de bloquant
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

        /// <summary>Vide le cache : la prochaine lecture de Utilisateurs ira relire le fichier JSON
        /// (pour récupérer les comptes modifiés depuis un autre poste, sans redémarrer).</summary>
        public static void InvaliderCacheUtilisateurs() => _cacheUtilisateurs = null;

        // -------- RÉSEAU : catalogue rubidiums --------

        private static List<Rubidium>? _cacheCatalogueRubidiums;

        /// <summary>
        /// Le catalogue des rubidiums prédéfinis, partagé entre les postes. Si le fichier
        /// réseau est vide ou absent, on renvoie un seed par défaut (Syref + Redondances)
        /// histoire d'amorcer l'UI avec quelque chose.
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

        /// <summary>Vide le cache : la prochaine lecture de CatalogueRubidiums ira relire le fichier réseau.</summary>
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

        /// <summary>Écriture atomique (on passe par un .tmp puis Replace) pour ne pas risquer de corrompre le fichier.</summary>
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
                // Si l'écriture réseau échoue (partage HS), la modification est perdue côté fichier.
                // Le cache mémoire, lui, garde l'état le temps de la session en cours.
            }
        }

        // -------- Migration vers le réseau --------

        /// <summary>One-shot : pousse vers le réseau les données qui traînent encore dans le settings.json local
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

        // -------- DTO de persistance --------

        private class Settings
        {
            public Rubidium? RubidiumActif { get; set; }
            public string? CheminMacroXLA { get; set; }

            // Champs legacy qu'on garde le temps de la migration one-shot vers le réseau.
            // Une fois migrés ils passent à null et ne sont plus sérialisés.
            public List<Rubidium>? CatalogueRubidiums { get; set; }
            public List<Utilisateur>? Utilisateurs { get; set; }
        }
    }
}
