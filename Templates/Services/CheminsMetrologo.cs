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
        /// Racine du partage serveur où sont centralisées toutes les données partagées
        /// entre les postes (Incertitudes, Presets, Catalogues, ArchivesLogs, etc.).
        /// Tous les dossiers réseau dérivent de cette racine.
        /// </summary>
        public const string BaseServeur = @"M:\exe_spe\Data_Metrologo";

        /// <summary>
        /// Chemin par défaut du fichier maître sur le serveur. Quand le poste n'a jamais
        /// été configuré, l'app le teste automatiquement au démarrage et l'enregistre comme
        /// source des chemins partagés.
        /// </summary>
        public static readonly string MasterPathsUrlDefaut =
            Path.Combine(BaseServeur, "paths.config.json");

        // Chemins serveur par défaut — utilisés pour amorcer le fichier master quand il
        // n'existe pas encore. Une fois écrits dans paths.config.json, les services lisent
        // ces emplacements via les propriétés Incertitudes/Presets/etc.
        public static readonly string IncertitudesServeurDefaut  = Path.Combine(BaseServeur, "Incertitudes");
        public static readonly string PresetsServeurDefaut       = Path.Combine(BaseServeur, "Presets");
        public static readonly string CataloguesServeurDefaut    = Path.Combine(BaseServeur, "Catalogues");
        public static readonly string ArchivesLogsServeurDefaut  = Path.Combine(BaseServeur, "ArchivesLogs");
        public static readonly string MesuresServeurDefaut       = Path.Combine(BaseServeur, "Mesures");
        public static readonly string UtilisateursServeurDefaut  = Path.Combine(BaseServeur, "Utilisateurs");
        public static readonly string RubidiumsServeurDefaut     = Path.Combine(BaseServeur, "Rubidiums");
        public static readonly string LogsServeurDefaut          = Path.Combine(BaseServeur, "Logs");
        public static readonly string BesanconServeurDefaut      = Path.Combine(BaseServeur, "Besancon");

        /// <summary>URL effective du fichier maître (depuis paths.config.json local).</summary>
        public static string MasterPathsUrl
        {
            get => _overrides.TryGetValue(CleMasterPathsUrl, out var v) ? v : string.Empty;
        }

        /// <summary>
        /// Charge les overrides de chemins. Stratégie :
        ///   1. Lit le cache local <c>Configuration\paths.config.json</c>.
        ///   2. Si aucun MasterPathsUrl n'y est défini, teste le défaut serveur
        ///      (<see cref="MasterPathsUrlDefaut"/>) — si le fichier existe, l'utilise
        ///      automatiquement (= bootstrap silencieux d'un poste vierge sans intervention admin).
        ///   3. Si un MasterPathsUrl est défini ET accessible → écrase les valeurs locales
        ///      par celles du maître ET met à jour le cache local pour usage hors-ligne.
        ///   4. Si le maître est inaccessible → garde le cache local = dernière config
        ///      connue valide, l'app continue en mode dégradé.
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

                // 2. Auto-pick : si pas de MasterPathsUrl configuré et que le défaut serveur
                // est accessible, on l'adopte automatiquement (poste vierge → bootstrap).
                if (string.IsNullOrWhiteSpace(MasterPathsUrl) && File.Exists(MasterPathsUrlDefaut))
                {
                    _overrides[CleMasterPathsUrl] = MasterPathsUrlDefaut;
                    try { SauverCacheLocal(); } catch { /* best-effort */ }
                }

                // 3. Récupère le maître depuis le serveur si URL renseignée
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
        /// Amorce la structure de dossiers sur le partage serveur si le partage est
        /// accessible et que la structure n'existe pas encore :
        ///   • <c>M:\exe_spe\Data_Metrologo\</c> + sous-dossiers (Incertitudes, Presets,
        ///     Catalogues, ArchivesLogs, Mesures)
        ///   • <c>M:\exe_spe\Data_Metrologo\paths.config.json</c> avec les chemins serveur
        ///     comme valeurs par défaut
        ///
        /// Idempotent : ne fait RIEN si tout est déjà en place. À appeler au démarrage de
        /// l'app sur tous les postes — le premier à avoir un partage accessible amorce la
        /// structure, les suivants la trouveront déjà prête.
        ///
        /// Retourne <c>true</c> si le partage est accessible (que la création ait été
        /// nécessaire ou non), <c>false</c> si M:\ n'est pas joignable.
        /// </summary>
        public static bool AssurerStructureServeur()
        {
            try
            {
                // Test d'accessibilité du partage : on crée la racine. Échoue silencieusement
                // si M:\ n'est pas monté (laptop en déplacement, serveur en maintenance...).
                Directory.CreateDirectory(BaseServeur);

                // Sous-dossiers de données partagées
                Directory.CreateDirectory(IncertitudesServeurDefaut);
                Directory.CreateDirectory(PresetsServeurDefaut);
                Directory.CreateDirectory(CataloguesServeurDefaut);
                Directory.CreateDirectory(ArchivesLogsServeurDefaut);
                Directory.CreateDirectory(MesuresServeurDefaut);
                Directory.CreateDirectory(UtilisateursServeurDefaut);
                Directory.CreateDirectory(RubidiumsServeurDefaut);
                Directory.CreateDirectory(LogsServeurDefaut);
                Directory.CreateDirectory(BesanconServeurDefaut);

                // Fichier maître paths.config.json — créé UNIQUEMENT s'il n'existe pas
                // (sinon on écraserait les modifs d'un autre admin).
                if (!File.Exists(MasterPathsUrlDefaut))
                {
                    var config = new Dictionary<string, string>
                    {
                        [nameof(Incertitudes)]  = IncertitudesServeurDefaut,
                        [nameof(Presets)]       = PresetsServeurDefaut,
                        [nameof(Catalogues)]    = CataloguesServeurDefaut,
                        [nameof(ArchivesLogs)]  = ArchivesLogsServeurDefaut,
                        [nameof(MesuresLocal)]  = MesuresServeurDefaut,
                        [nameof(Utilisateurs)]  = UtilisateursServeurDefaut,
                        [nameof(Rubidiums)]     = RubidiumsServeurDefaut,
                        [nameof(Logs)]          = LogsServeurDefaut,
                        [nameof(Besancon)]      = BesanconServeurDefaut,
                    };
                    string json = JsonSerializer.Serialize(config,
                        new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(MasterPathsUrlDefaut, json);
                }
                else
                {
                    // Le master existe déjà : on ajoute les clés manquantes (migration
                    // douce vers les nouveaux chemins partagés Utilisateurs/Rubidiums/Logs
                    // sans toucher aux clés déjà configurées par l'admin).
                    try
                    {
                        string existant = File.ReadAllText(MasterPathsUrlDefaut);
                        var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(existant)
                                   ?? new Dictionary<string, string>();
                        bool modifie = false;
                        void AjouterSiAbsent(string cle, string valeur)
                        {
                            if (!dict.ContainsKey(cle))
                            {
                                dict[cle] = valeur;
                                modifie = true;
                            }
                        }
                        AjouterSiAbsent(nameof(Utilisateurs), UtilisateursServeurDefaut);
                        AjouterSiAbsent(nameof(Rubidiums), RubidiumsServeurDefaut);
                        AjouterSiAbsent(nameof(Logs), LogsServeurDefaut);
                        AjouterSiAbsent(nameof(Besancon), BesanconServeurDefaut);
                        if (modifie)
                        {
                            string json = JsonSerializer.Serialize(dict,
                                new JsonSerializerOptions { WriteIndented = true });
                            File.WriteAllText(MasterPathsUrlDefaut, json);
                        }
                    }
                    catch { /* best-effort */ }
                }
                return true;
            }
            catch
            {
                // Partage serveur indisponible → on retombe gracieusement sur les chemins
                // locaux par défaut. L'app reste utilisable en local seul.
                return false;
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

        /// <summary>
        /// Dossier des profils d'appareils legacy (EIP / Racal / Stanford) — commandes GPIB
        /// éditables à la main. Sur le réseau par défaut
        /// (<c>M:\exe_spe\Data_Metrologo\Appareils_Legacy</c>), surchargeable (clé <c>AppareilsLegacy</c>).
        /// </summary>
        public static string AppareilsLegacy =>
            AvecOverride(nameof(AppareilsLegacy), Path.Combine(BaseServeur, "Appareils_Legacy"));

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

        /// <summary>Utilisateurs partagés entre postes — surchargeable réseau (clé <c>Utilisateurs</c>).</summary>
        public static string Utilisateurs =>
            AvecOverride(nameof(Utilisateurs), Path.Combine(Racine, "Utilisateurs"));

        /// <summary>Catalogue rubidiums partagé — surchargeable réseau (clé <c>Rubidiums</c>).</summary>
        public static string Rubidiums =>
            AvecOverride(nameof(Rubidiums), Path.Combine(Racine, "Rubidiums"));

        /// <summary>Logs courants partagés entre postes — surchargeable réseau (clé <c>Logs</c>).</summary>
        public static string Logs =>
            AvecOverride(nameof(Logs), Path.Combine(Racine, "Logs"));

        /// <summary>Données Besançon (valeurs journalières + moyennes hebdo) partagées entre postes
        /// — surchargeable réseau (clé <c>Besancon</c>) → M:\exe_spe\Data_Metrologo\Besancon.</summary>
        public static string Besancon =>
            AvecOverride(nameof(Besancon), Path.Combine(Racine, "Besancon"));

        /// <summary>Fichier JSON des utilisateurs + hash mdp.</summary>
        public static string FichierUtilisateurs =>
            Path.Combine(Utilisateurs, "utilisateurs.json");

        /// <summary>Fichier JSON du catalogue rubidiums partagé.</summary>
        public static string FichierCatalogueRubidiums =>
            Path.Combine(Rubidiums, "catalogue.json");

        /// <summary>Fichier JSON du rubidium ACTIF partagé entre tous les postes (référence
        /// commune). Lu au démarrage pour que chaque poste reprenne le rubidium sélectionné.</summary>
        public static string FichierRubidiumActif =>
            Path.Combine(Rubidiums, "rubidium-actif.json");

        /// <summary>Fichier JSON-lines des sessions de journal.</summary>
        public static string FichierJournalSessions =>
            Path.Combine(Logs, "sessions.jsonl");

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
        /// <c>&lt;MesuresLocal&gt;\&lt;FI&gt;\Mesures_*.xlsx</c> juste après la sauvegarde
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

        /// <summary>
        /// Fichier JSON du catalogue d'appareils partagé entre postes — vit dans le
        /// dossier <see cref="Catalogues"/> (surchargeable réseau par défaut sur M:\).
        /// Remplace l'ancien stockage SQL Server.
        /// </summary>
        public static string FichierCatalogueAppareils =>
            Path.Combine(Catalogues, "appareils.json");

        /// <summary>
        /// Fichier JSON des profils d'appareils legacy (EIP / Racal / Stanford). Créé au 1er
        /// lancement avec les valeurs par défaut, puis relu à chaque démarrage : permet de
        /// corriger les commandes GPIB sur le réseau sans recompiler l'application.
        /// </summary>
        public static string FichierAppareilsLegacy =>
            Path.Combine(AppareilsLegacy, "appareils-legacy.json");

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
