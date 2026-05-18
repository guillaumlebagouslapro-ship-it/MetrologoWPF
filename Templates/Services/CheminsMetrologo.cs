using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace Metrologo.Services
{
    /// <summary>
    /// Source unique des chemins de fichiers utilisés par l'application en local
    /// (<c>%LocalAppData%\Metrologo\</c>). Tout service qui lit/écrit un fichier
    /// utilisateur passe par cette classe — évite la duplication de chemins
    /// hardcodés à droite/gauche et permet une réorganisation par sous-dossiers
    /// claire et cohérente.
    /// <para/>
    /// La méthode <see cref="MigrerAnciensFichiers"/> est appelée au démarrage de
    /// l'app (cf. <c>App.OnStartup</c>) : elle déplace silencieusement les fichiers
    /// historiquement à plat dans <c>%LocalAppData%\Metrologo\</c> vers les nouveaux
    /// sous-dossiers (<c>Configuration\</c>, <c>Presets\</c>, etc.). Idempotente,
    /// safe à rappeler à chaque démarrage.
    /// </summary>
    public static class CheminsMetrologo
    {
        // ---------- Overrides chargés depuis paths.config.json ----------

        private static Dictionary<string, string> _overrides = new();

        /// <summary>
        /// Charge les overrides de chemins depuis <c>Configuration\paths.config.json</c>.
        /// Si le fichier existe, les chemins définis dedans surchargent les défauts locaux —
        /// permet de pointer Incertitudes / Presets / Catalogues vers un partage réseau
        /// commun à tous les postes du site sans toucher au code.
        /// Idempotent : safe à appeler à chaque démarrage.
        /// </summary>
        public static void ChargerConfigChemins()
        {
            try
            {
                string fichier = FichierPathsConfig;
                if (!File.Exists(fichier)) { _overrides.Clear(); return; }

                string json = File.ReadAllText(fichier);
                var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                _overrides = dict ?? new Dictionary<string, string>();
            }
            catch
            {
                // JSON corrompu / inaccessible → fallback aux chemins par défaut.
                _overrides = new Dictionary<string, string>();
            }
        }

        /// <summary>
        /// Sauvegarde les overrides dans <c>Configuration\paths.config.json</c>. Champs vides
        /// = on retire l'override (= retombe sur le défaut local).
        /// </summary>
        public static void EnregistrerConfigChemins(Dictionary<string, string> overrides)
        {
            // Nettoie les champs vides — un chemin vide = pas d'override = défaut local.
            var aGarder = new Dictionary<string, string>();
            foreach (var kv in overrides)
            {
                if (!string.IsNullOrWhiteSpace(kv.Value))
                    aGarder[kv.Key] = kv.Value.Trim();
            }
            _overrides = aGarder;

            Directory.CreateDirectory(Configuration);
            string json = JsonSerializer.Serialize(_overrides,
                new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(FichierPathsConfig, json);
        }

        /// <summary>Indique si un chemin a été surchargé par paths.config.json.</summary>
        public static bool EstSurcharge(string nomCle) =>
            _overrides.TryGetValue(nomCle, out var c) && !string.IsNullOrWhiteSpace(c);

        /// <summary>Retourne l'override s'il existe, sinon le chemin par défaut fourni.</summary>
        private static string AvecOverride(string nomCle, string parDefaut) =>
            _overrides.TryGetValue(nomCle, out var c) && !string.IsNullOrWhiteSpace(c)
                ? c
                : parDefaut;

        // ---------- Dossiers ----------

        /// <summary>Racine locale (jamais surchargée — toujours <c>%LocalAppData%\Metrologo\</c>).</summary>
        public static string Racine => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Metrologo");

        /// <summary>Settings UI, configuration globale, credentials BDD. Toujours local
        /// (par poste, contient des secrets DPAPI liés au profil Windows).</summary>
        public static string Configuration => Path.Combine(Racine, "Configuration");

        /// <summary>Backups d'imports — surchargeable réseau (clé <c>Catalogues</c>).</summary>
        public static string Catalogues =>
            AvecOverride(nameof(Catalogues), Path.Combine(Racine, "Catalogues"));

        /// <summary>Presets stabilité — surchargeable réseau (clé <c>Presets</c>).</summary>
        public static string Presets =>
            AvecOverride(nameof(Presets), Path.Combine(Racine, "Presets"));

        /// <summary>Caches locaux. Toujours local (cache, ne fait pas sens en réseau).</summary>
        public static string Cache => Path.Combine(Racine, "Cache");

        /// <summary>Modules d'incertitude — surchargeable réseau (clé <c>Incertitudes</c>).</summary>
        public static string Incertitudes =>
            AvecOverride(nameof(Incertitudes), Path.Combine(Racine, "Incertitudes"));

        /// <summary>Archives logs — surchargeable réseau (clé <c>ArchivesLogs</c>).</summary>
        public static string ArchivesLogs =>
            AvecOverride(nameof(ArchivesLogs), Path.Combine(Racine, "ArchivesLogs"));

        /// <summary>
        /// Chemin local de duplication des fichiers Excel produits par chaque mesure.
        /// Vide par défaut (= pas de duplication). Une fois configuré, chaque rapport
        /// est copié vers <c>&lt;CheminMesuresLocal&gt;\&lt;FI&gt;\Mesures_*.xlsm</c> juste
        /// après la sauvegarde principale — garantit qu'une coupure réseau avec le
        /// serveur principal n'entraîne pas la perte du rapport (copie locale dispo).
        /// Surchargeable via paths.config.json (clé <c>MesuresLocal</c>).
        /// </summary>
        public static string MesuresLocal =>
            AvecOverride(nameof(MesuresLocal), string.Empty);

        /// <summary>Vrai si l'utilisateur a configuré un chemin local de duplication.</summary>
        public static bool MesuresLocalConfigure =>
            !string.IsNullOrWhiteSpace(MesuresLocal);

        /// <summary>Fichier de configuration des chemins. Toujours local.</summary>
        public static string FichierPathsConfig =>
            Path.Combine(Configuration, "paths.config.json");

        // ---------- Fichiers ----------

        public static string FichierSettings        => Path.Combine(Configuration, "settings.json");
        public static string FichierMesuresConfig   => Path.Combine(Configuration, "Mesures_Config.txt");
        public static string FichierDbCredentials   => Path.Combine(Configuration, "db.credentials");

        public static string FichierPresetsStabilite => Path.Combine(Presets, "PresetsStabilite.json");

        public static string FichierJournalDbLegacy  => Path.Combine(Cache, "journal.db");

        public static string FichierAppareilsCatalogueLegacy =>
            Path.Combine(Catalogues, "AppareilsCatalogue.json");

        // ---------- Migration ----------

        /// <summary>
        /// Crée les sous-dossiers + déplace les fichiers historiquement à plat dans
        /// <c>%LocalAppData%\Metrologo\</c> vers leur nouveau sous-dossier dédié.
        /// Idempotent : ne fait rien si les fichiers sont déjà à leur place finale.
        /// Erreurs swallowed (best-effort) : un fichier qui ne se déplace pas n'empêche
        /// pas le démarrage.
        /// </summary>
        public static void MigrerAnciensFichiers()
        {
            try
            {
                // Crée tous les sous-dossiers (Directory.CreateDirectory est idempotent).
                Directory.CreateDirectory(Racine);
                Directory.CreateDirectory(Configuration);
                Directory.CreateDirectory(Catalogues);
                Directory.CreateDirectory(Presets);
                Directory.CreateDirectory(Cache);
                Directory.CreateDirectory(Incertitudes);

                // Fichiers individuels à déplacer (ancien nom à plat → nouveau chemin).
                DeplacerSiPresent("settings.json",          FichierSettings);
                DeplacerSiPresent("Mesures_Config.txt",     FichierMesuresConfig);
                DeplacerSiPresent("db.credentials",         FichierDbCredentials);
                DeplacerSiPresent("PresetsStabilite.json",  FichierPresetsStabilite);
                DeplacerSiPresent("journal.db",             FichierJournalDbLegacy);
                DeplacerSiPresent("journal.db-shm",         FichierJournalDbLegacy + "-shm");
                DeplacerSiPresent("journal.db-wal",         FichierJournalDbLegacy + "-wal");

                // Backups d'import du catalogue (motif "AppareilsCatalogue.json.imported.*"
                // créés par l'ancienne migration JSON → SQL Server) → Catalogues\.
                foreach (var f in Directory.GetFiles(Racine, "AppareilsCatalogue.json*"))
                {
                    string nom = Path.GetFileName(f);
                    string dest = Path.Combine(Catalogues, nom);
                    DeplacerFichier(f, dest);
                }
            }
            catch
            {
                // Best-effort : un échec de migration ne doit pas empêcher le démarrage
                // (les services retomberont sur les anciens chemins via fallback ci-dessous).
            }
        }

        /// <summary>
        /// Cherche un fichier d'abord à son nouveau chemin, sinon à l'ancien chemin (à plat
        /// dans Racine). Utilisé par les services pour rester rétro-compatible avec une
        /// installation où la migration n'aurait pas encore tourné (ex: appel à un service
        /// avant <see cref="MigrerAnciensFichiers"/>). Retourne le 1er chemin existant, ou
        /// <paramref name="nouveauChemin"/> si aucun n'existe (= chemin de création).
        /// </summary>
        public static string ResoudreCheminAvecFallback(string nouveauChemin, string ancienNomRelatif)
        {
            if (File.Exists(nouveauChemin)) return nouveauChemin;
            string ancien = Path.Combine(Racine, ancienNomRelatif);
            if (File.Exists(ancien)) return ancien;
            return nouveauChemin;
        }

        // ---------- Internes ----------

        private static void DeplacerSiPresent(string ancienNomRelatif, string nouveauChemin)
        {
            string ancien = Path.Combine(Racine, ancienNomRelatif);
            DeplacerFichier(ancien, nouveauChemin);
        }

        private static void DeplacerFichier(string ancien, string nouveau)
        {
            try
            {
                if (!File.Exists(ancien)) return;
                if (File.Exists(nouveau)) return;   // déjà migré, ne pas écraser
                Directory.CreateDirectory(Path.GetDirectoryName(nouveau)!);
                File.Move(ancien, nouveau);
            }
            catch
            {
                /* best-effort silencieux — fallback gérera */
            }
        }
    }
}
