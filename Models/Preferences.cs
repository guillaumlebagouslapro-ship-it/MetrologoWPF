using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Metrologo.Services;

namespace Metrologo.Models
{
    /// <summary>
    /// Préférences de l'application réparties sur deux couches :
    ///
    ///   • LOCAL — fichier <c>%LocalAppData%\Metrologo\Configuration\settings.json</c>.
    ///     Contient ce qui est spécifique au poste : rubidium actif courant + chemin
    ///     de la macro Excel personnalisée.
    ///
    ///   • RÉSEAU — fichiers dédiés sur le partage M:\ pour ce qui doit être partagé
    ///     entre tous les postes :
    ///       <c>Utilisateurs\utilisateurs.json</c> (liste des comptes + hash mdp)
    ///       <c>Rubidiums\catalogue.json</c>     (catalogue des rubidiums dispo)
    ///
    /// Chaque fichier réseau est ré-lu à chaque accès aux propriétés concernées —
    /// ce qui permet à un poste de voir immédiatement les ajouts/modifs faits par
    /// un autre, sans redémarrage. Cache mémoire avec invalidation au write.
    /// </summary>
    public static class Preferences
    {
        private static Settings _settings = new();

        // -------- LOCAL : rubidium actif + macro --------

        /// <summary>
        /// Rubidium actif. Lu en priorité depuis le fichier PARTAGÉ
        /// (<see cref="CheminsMetrologo.FichierRubidiumActif"/>) pour que tous les postes
        /// reprennent le même rubidium de référence (au démarrage). Repli sur le réglage local
        /// si le partage est absent/injoignable.
        /// </summary>
        public static Rubidium? RubidiumActif =>
            LireFichierReseau<Rubidium>(CheminsMetrologo.FichierRubidiumActif) ?? _settings.RubidiumActif;

        public static string CheminMacroXLA
        {
            get => _settings.CheminMacroXLA ?? @"C:\Exe_Spe\Fct_VBA\Metrologo.xla";
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
                // Silencieux — la perte de préférence locale n'est pas bloquante.
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

        /// <summary>
        /// Invalide le cache mémoire des utilisateurs : la prochaine lecture de
        /// <see cref="Utilisateurs"/> relira le fichier JSON. Permet de refléter en temps réel
        /// les comptes ajoutés / modifiés depuis un autre poste (ou via la gestion des
        /// utilisateurs) sans redémarrer l'application.
        /// </summary>
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

        /// <summary>
        /// One-shot : si <c>settings.json</c> local contient encore des données qui
        /// ont été déplacées vers le réseau (utilisateurs, catalogue rubidiums), on
        /// migre ces données vers les fichiers réseau puis on retire les champs du
        /// settings local. Idempotent : ne fait rien si les fichiers réseau existent
        /// déjà ou si les champs locaux sont vides.
        /// </summary>
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
