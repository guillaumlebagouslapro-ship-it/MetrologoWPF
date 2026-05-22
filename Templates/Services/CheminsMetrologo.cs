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
        /// Clé spéciale dans paths.config.json (locale ET master) qui pointe vers le fichier
        /// MAÎTRE sur le partage serveur. Quand cette clé est renseignée, l'app lit le fichier
        /// distant au démarrage et écrase les valeurs locales — permet à l'admin de modifier
        /// les chemins UNE seule fois (sur le serveur) et de propager à TOUS les postes
        /// automatiquement, sans intervention manuelle sur chaque PC.
        /// </summary>
        public const string CleMasterPathsUrl = "_MasterPathsUrl";

        /// <summary>
        /// Chemin par défaut du fichier maître sur le serveur (utilisé pour pré-remplir
        /// l'UI Admin si le poste n'a encore aucune config). L'admin peut le changer.
        /// </summary>
        public const string MasterPathsUrlDefaut = @"M:\exe_spe\Data_Metrologo\paths.config.json";

        /// <summary>URL effective du fichier maître (depuis paths.config.json local).</summary>
        public static string MasterPathsUrl
        {
            get => _overrides.TryGetValue(CleMasterPathsUrl, out var v) ? v : string.Empty;
        }

        /// <summary>
        /// Charge les overrides de chemins. Stratégie 2 niveaux :
        ///   1. Lit le fichier local <c>Configuration\paths.config.json</c> (cache).
        ///   2. Si la clé <c>_MasterPathsUrl</c> y est définie ET que le fichier maître
        ///      est accessible → écrase les valeurs locales par celles du maître ET
        ///      met à jour le cache local pour usage hors-ligne.
        ///   3. Si le maître est inaccessible (réseau down, serveur en maintenance) →
        ///      garde les valeurs locales = dernière config connue valide. L'app continue
        ///      à fonctionner en mode dégradé.
        /// Idempotent : safe à appeler à chaque démarrage.
        /// </summary>
        public static void ChargerConfigChemins()
        {
            try
            {
                // 1. Charge le cache local
                string fichier = FichierPathsConfig;
                if (File.Exists(fichier))
                {
                    string json = File.ReadAllText(fichier);
                    var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                    _overrides = dict ?? new Dictionary<string, string>();
                }
                else
                {
                    _overrides.Clear();
                }

                // 2. Tente de récupérer le maître depuis le serveur si URL renseignée
                string masterUrl = MasterPathsUrl;
                if (!string.IsNullOrWhiteSpace(masterUrl) && File.Exists(masterUrl))
                {
                    try
                    {
                        string jsonMaster = File.ReadAllText(masterUrl);
                        var dictMaster = JsonSerializer.Deserialize<Dictionary<string, string>>(jsonMaster);
                        if (dictMaster != null)
                        {
                            // Master gagne. On préserve la _MasterPathsUrl locale (sinon
                            // un master sans cette clé briserait le mécanisme au prochain run).
                            foreach (var kv in dictMaster)
                            {
                                if (!string.Equals(kv.Key, CleMasterPathsUrl, StringComparison.OrdinalIgnoreCase))
                                    _overrides[kv.Key] = kv.Value;
                            }
                            // Mise à jour du cache local pour usage hors-ligne au prochain démarrage.
                            try { SauverCacheLocal(); } catch { /* best-effort */ }
                        }
                    }
                    catch
                    {
                        // Master inaccessible / corrompu → on garde le cache local tel quel.
                        // Pas de log fatal — le mode dégradé est intentionnel.
                    }
                }
            }
            catch
            {
                // JSON corrompu / inaccessible → fallback aux chemins par défaut.
                _overrides = new Dictionary<string, string>();
            }
        }

        /// <summary>
        /// Sauvegarde les overrides. Toujours dans le cache local. Si
        /// <paramref name="appliquerATousLesPostes"/> est vrai ET qu'un MasterPathsUrl est
        /// renseigné dans les overrides, écrit AUSSI le fichier maître sur le serveur —
        /// tous les autres postes prendront ces nouvelles valeurs à leur prochain démarrage.
        /// Champs vides = on retire l'override (= retombe sur le défaut local).
        /// </summary>
        public static void EnregistrerConfigChemins(Dictionary<string, string> overrides,
                                                    bool appliquerATousLesPostes = false)
        {
            // Nettoie les champs vides — un chemin vide = pas d'override = défaut local.
            var aGarder = new Dictionary<string, string>();
            foreach (var kv in overrides)
            {
                if (!string.IsNullOrWhiteSpace(kv.Value))
                    aGarder[kv.Key] = kv.Value.Trim();
            }
            _overrides = aGarder;

            SauverCacheLocal();

            // Propagation vers le maître si demandé
            if (appliquerATousLesPostes && _overrides.TryGetValue(CleMasterPathsUrl, out var masterUrl)
                && !string.IsNullOrWhiteSpace(masterUrl))
            {
                try
                {
                    string? dossierMaster = Path.GetDirectoryName(masterUrl);
                    if (!string.IsNullOrEmpty(dossierMaster))
                        Directory.CreateDirectory(dossierMaster);

                    // Le master ne contient PAS la clé _MasterPathsUrl (chaque poste a la sienne)
                    var aEcrireMaster = new Dictionary<string, string>(_overrides);
                    aEcrireMaster.Remove(CleMasterPathsUrl);
                    string jsonMaster = JsonSerializer.Serialize(aEcrireMaster,
                        new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(masterUrl, jsonMaster);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        $"Le cache local a été sauvegardé mais le fichier maître "
                      + $"« {masterUrl} » n'a pas pu être écrit : {ex.Message}", ex);
                }
            }
        }

        /// <summary>Sauvegarde uniquement le cache local (interne).</summary>
        private static void SauverCacheLocal()
        {
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
        /// Chemin par défaut (générique, identique sur tous les postes) où sont dupliqués
        /// les rapports Excel des mesures. Pointe sur le dossier Documents public Windows
        /// (<c>C:\Users\Public\Documents\Metrologo_Backup</c>) — accessible en écriture par
        /// tous les utilisateurs du poste sans privilèges admin requis, contrairement à
        /// la racine <c>C:\</c>. Créé automatiquement au démarrage de l'app si absent
        /// (cf. App.OnStartup → AssurerDossierMesuresLocal).
        /// </summary>
        public static string MesuresLocalDefaut => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments),
            "Metrologo_Backup");

        /// <summary>
        /// Chemin local de duplication des fichiers Excel produits par chaque mesure.
        /// Par défaut <see cref="MesuresLocalDefaut"/> (commun à tous les postes). Surchargeable
        /// via paths.config.json (clé <c>MesuresLocal</c>) si un admin veut le pointer ailleurs
        /// pour un poste spécifique. Chaque rapport est copié vers
        /// <c>&lt;MesuresLocal&gt;\&lt;FI&gt;\Mesures_*.xlsm</c> juste après la sauvegarde
        /// principale — garantit qu'une coupure réseau avec le serveur principal n'entraîne
        /// pas la perte du rapport (copie locale dispo).
        /// </summary>
        public static string MesuresLocal =>
            AvecOverride(nameof(MesuresLocal), MesuresLocalDefaut);

        /// <summary>
        /// Vrai si un chemin local non vide est défini. Avec le défaut générique, c'est
        /// toujours vrai sauf si un admin a explicitement mis le champ à vide.
        /// </summary>
        public static bool MesuresLocalConfigure =>
            !string.IsNullOrWhiteSpace(MesuresLocal);

        /// <summary>
        /// S'assure que le dossier <see cref="MesuresLocal"/> existe sur le poste, le crée
        /// sinon. Idempotent. Best-effort : un échec (permissions, disque inaccessible) est
        /// loggué mais ne lève pas — la duplication des mesures retombera silencieusement
        /// en cas d'échec d'écriture ultérieur.
        /// </summary>
        public static bool AssurerDossierMesuresLocal()
        {
            string chemin = MesuresLocal;
            if (string.IsNullOrWhiteSpace(chemin)) return false;
            try
            {
                Directory.CreateDirectory(chemin);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Flag indiquant que le premier démarrage de l'app sur ce poste a déjà été traité
        /// (message d'accueil + raccourci bureau créés). Présent dans Configuration\ pour
        /// rester par-poste (n'est pas synchronisé sur le réseau).
        /// </summary>
        public static string FichierFirstRunFlag =>
            Path.Combine(Configuration, "first_run_done.flag");

        /// <summary>Vrai au tout premier démarrage de l'app sur ce poste.</summary>
        public static bool EstPremierDemarrage() => !File.Exists(FichierFirstRunFlag);

        /// <summary>Marque le premier démarrage comme effectué (créé le flag).</summary>
        public static void MarquerPremierDemarrageEffectue()
        {
            try
            {
                Directory.CreateDirectory(Configuration);
                File.WriteAllText(FichierFirstRunFlag,
                    $"Premier démarrage effectué le {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            }
            catch { /* best-effort */ }
        }

        /// <summary>
        /// Crée un raccourci .lnk sur le Bureau de l'utilisateur courant pointant vers le
        /// dossier local de sauvegarde des mesures. Distinct du raccourci serveur par le
        /// suffixe « (local) » dans son nom. Idempotent : si le raccourci existe déjà,
        /// ne le recrée pas. Retourne le chemin du raccourci créé, ou null si échec.
        /// </summary>
        public static string? CreerRaccourciBureauMesuresLocal()
        {
            try
            {
                string bureau = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string cheminRaccourci = Path.Combine(bureau, "Mesures Metrologo (local).lnk");

                if (File.Exists(cheminRaccourci)) return cheminRaccourci;

                // Création via COM Wscript.Shell — pas de dépendance externe nécessaire,
                // dispo nativement sur Windows.
                var t = Type.GetTypeFromProgID("WScript.Shell");
                if (t == null) return null;
                dynamic shell = Activator.CreateInstance(t)!;
                dynamic raccourci = shell.CreateShortcut(cheminRaccourci);
                raccourci.TargetPath  = MesuresLocal;
                raccourci.Description = "Dossier local de sauvegarde des mesures Metrologo";
                raccourci.WorkingDirectory = MesuresLocal;
                raccourci.Save();
                return cheminRaccourci;
            }
            catch
            {
                return null;
            }
        }

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
