using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace Metrologo.Services
{
    /// <summary>
    /// Le point d'entrée unique pour tous les chemins de fichiers que l'application
    /// manipule en local (<c>%LocalAppData%\Metrologo\</c>). Dès qu'un service a besoin
    /// de lire ou écrire un fichier utilisateur, il passe par ici. Comme ça, on ne se
    /// retrouve pas avec des chemins en dur éparpillés partout, et on peut réorganiser
    /// les sous-dossiers proprement et au même endroit.
    /// <para/>
    /// Au démarrage de l'app (voir <c>App.OnStartup</c>), on appelle
    /// <see cref="MigrerAnciensFichiers"/> : elle déplace en douce les fichiers qui
    /// traînaient à plat dans <c>%LocalAppData%\Metrologo\</c> vers leurs nouveaux
    /// sous-dossiers (<c>Configuration\</c>, <c>Presets\</c>, etc.). Elle est idempotente,
    /// donc aucun souci à la rappeler à chaque lancement.
    /// </summary>
    public static class CheminsMetrologo
    {
        // ---------- Overrides lus depuis paths.config.json ----------

        private static Dictionary<string, string> _overrides = new();

        /// <summary>
        /// Clé un peu particulière dans paths.config.json (présente côté local comme côté
        /// master) : elle pointe vers le fichier MAÎTRE sur le partage serveur. Si elle est
        /// renseignée, l'app va lire ce fichier distant au démarrage et ses valeurs prennent
        /// le pas sur le local. L'intérêt : l'admin ne modifie les chemins qu'une seule fois,
        /// sur le serveur, et le changement se propage tout seul à TOUS les postes, sans avoir
        /// à passer sur chaque PC à la main.
        /// </summary>
        public const string CleMasterPathsUrl = "_MasterPathsUrl";

        /// <summary>
        /// La racine du partage serveur, là où on centralise tout ce que les postes
        /// partagent (Incertitudes, Presets, Catalogues, ArchivesLogs, etc.). Tous les
        /// dossiers réseau partent de cette racine.
        /// </summary>
        public const string BaseServeur = @"M:\exe_spe\Data_Metrologo";

        /// <summary>
        /// L'emplacement par défaut du fichier maître sur le serveur. Sur un poste qui n'a
        /// jamais été configuré, l'app va tester ce chemin au démarrage et, s'il répond,
        /// l'adopter comme source des chemins partagés.
        /// </summary>
        public static readonly string MasterPathsUrlDefaut =
            Path.Combine(BaseServeur, "paths.config.json");

        // Chemins serveur par défaut. Ils servent à amorcer le fichier master la première
        // fois, quand il n'existe pas encore. Une fois écrits dans paths.config.json, ce
        // sont les propriétés Incertitudes/Presets/etc. qui les exposent aux services.
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
        /// Charge les overrides de chemins. Voici comment ça se déroule :
        ///   1. On lit le cache local <c>Configuration\paths.config.json</c>.
        ///   2. S'il n'y a pas de MasterPathsUrl dedans, on tente le défaut serveur
        ///      (<see cref="MasterPathsUrlDefaut"/>). Si le fichier répond, on l'adopte
        ///      directement : un poste tout neuf se configure ainsi tout seul, sans que
        ///      l'admin ait à intervenir.
        ///   3. Si un MasterPathsUrl est défini et qu'on arrive à l'atteindre, ses valeurs
        ///      remplacent les locales et on en profite pour rafraîchir le cache local
        ///      (utile pour le mode hors-ligne).
        ///   4. Si le maître est injoignable, on s'en tient au cache local, c'est-à-dire à
        ///      la dernière config valide connue, et l'app continue en mode dégradé.
        /// La méthode est idempotente : aucun problème à l'appeler à chaque démarrage.
        /// </summary>
        public static void ChargerConfigChemins()
        {
            try
            {
                // 1. On commence par charger le cache local
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

                // 2. Choix automatique : pas de MasterPathsUrl configuré mais le défaut
                // serveur répond ? On l'adopte direct (cas du poste vierge qui s'amorce).
                if (string.IsNullOrWhiteSpace(MasterPathsUrl) && File.Exists(MasterPathsUrlDefaut))
                {
                    _overrides[CleMasterPathsUrl] = MasterPathsUrlDefaut;
                    try { SauverCacheLocal(); } catch { /* best-effort */ }
                }

                // 3. On va chercher le maître sur le serveur si une URL est renseignée
                string masterUrl = MasterPathsUrl;
                if (!string.IsNullOrWhiteSpace(masterUrl) && File.Exists(masterUrl))
                {
                    try
                    {
                        string jsonMaster = File.ReadAllText(masterUrl);
                        var dictMaster = JsonSerializer.Deserialize<Dictionary<string, string>>(jsonMaster);
                        if (dictMaster != null)
                        {
                            // Le master a le dernier mot, mais on garde quand même la
                            // _MasterPathsUrl locale : sans elle, un master qui ne la
                            // contient pas casserait tout le mécanisme au prochain run.
                            foreach (var kv in dictMaster)
                            {
                                if (!string.Equals(kv.Key, CleMasterPathsUrl, StringComparison.OrdinalIgnoreCase))
                                    _overrides[kv.Key] = kv.Value;
                            }
                            // On rafraîchit le cache local, ça servira au prochain démarrage hors-ligne.
                            try { SauverCacheLocal(); } catch { /* best-effort */ }
                        }
                    }
                    catch
                    {
                        // Master injoignable ou corrompu : on laisse le cache local tel
                        // quel. Pas de log fatal, le mode dégradé est voulu.
                    }
                }
            }
            catch
            {
                // JSON corrompu ou illisible : on retombe sur les chemins par défaut.
                _overrides = new Dictionary<string, string>();
            }
        }

        /// <summary>
        /// Met en place la structure de dossiers sur le partage serveur, à condition que le
        /// partage soit accessible et que la structure n'existe pas déjà :
        ///   • <c>M:\exe_spe\Data_Metrologo\</c> et ses sous-dossiers (Incertitudes, Presets,
        ///     Catalogues, ArchivesLogs, Mesures)
        ///   • <c>M:\exe_spe\Data_Metrologo\paths.config.json</c>, rempli avec les chemins
        ///     serveur comme valeurs par défaut
        ///
        /// Idempotente : si tout est déjà en place, elle ne touche à rien. On l'appelle au
        /// démarrage sur tous les postes ; le premier qui a un partage accessible crée la
        /// structure, et les autres la trouveront déjà prête.
        ///
        /// Renvoie <c>true</c> dès que le partage est accessible (qu'il ait fallu créer
        /// quelque chose ou non), et <c>false</c> si M:\ ne répond pas.
        /// </summary>
        public static bool AssurerStructureServeur()
        {
            try
            {
                // Pour tester si le partage répond, on tente simplement de créer la racine.
                // Ça échoue sans bruit si M:\ n'est pas monté (portable en déplacement,
                // serveur en maintenance, etc.).
                Directory.CreateDirectory(BaseServeur);

                // Les sous-dossiers de données partagées
                Directory.CreateDirectory(IncertitudesServeurDefaut);
                Directory.CreateDirectory(PresetsServeurDefaut);
                Directory.CreateDirectory(CataloguesServeurDefaut);
                Directory.CreateDirectory(ArchivesLogsServeurDefaut);
                Directory.CreateDirectory(MesuresServeurDefaut);
                Directory.CreateDirectory(UtilisateursServeurDefaut);
                Directory.CreateDirectory(RubidiumsServeurDefaut);
                Directory.CreateDirectory(LogsServeurDefaut);
                Directory.CreateDirectory(BesanconServeurDefaut);

                // Le fichier maître paths.config.json : on ne le crée que s'il n'existe
                // pas, sinon on écraserait les réglages d'un autre admin.
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
                    // Le master est déjà là : on se contente d'ajouter les clés qui
                    // manquent. C'est une migration en douceur vers les nouveaux chemins
                    // partagés (Utilisateurs/Rubidiums/Logs), sans toucher à ce que l'admin
                    // a déjà configuré.
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
                // Partage serveur indisponible : on bascule proprement sur les chemins
                // locaux par défaut, l'app reste tout à fait utilisable en local seul.
                return false;
            }
        }

        /// <summary>
        /// Enregistre les overrides. Le cache local est toujours mis à jour. En plus, si
        /// <paramref name="appliquerATousLesPostes"/> est vrai et qu'un MasterPathsUrl figure
        /// dans les overrides, on écrit aussi le fichier maître sur le serveur : du coup, tous
        /// les autres postes récupéreront ces valeurs à leur prochain démarrage.
        /// Un champ laissé vide veut dire qu'on supprime l'override, et donc qu'on revient au
        /// défaut local.
        /// </summary>
        public static void EnregistrerConfigChemins(Dictionary<string, string> overrides,
                                                    bool appliquerATousLesPostes = false)
        {
            // On écarte les champs vides : un chemin vide signifie pas d'override, donc défaut local.
            var aGarder = new Dictionary<string, string>();
            foreach (var kv in overrides)
            {
                if (!string.IsNullOrWhiteSpace(kv.Value))
                    aGarder[kv.Key] = kv.Value.Trim();
            }
            _overrides = aGarder;

            SauverCacheLocal();

            // On propage vers le maître si on nous l'a demandé
            if (appliquerATousLesPostes && _overrides.TryGetValue(CleMasterPathsUrl, out var masterUrl)
                && !string.IsNullOrWhiteSpace(masterUrl))
            {
                try
                {
                    string? dossierMaster = Path.GetDirectoryName(masterUrl);
                    if (!string.IsNullOrEmpty(dossierMaster))
                        Directory.CreateDirectory(dossierMaster);

                    // Le master ne porte pas la clé _MasterPathsUrl : chaque poste a la sienne
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

        /// <summary>La racine locale. On ne la surcharge jamais : c'est toujours <c>%LocalAppData%\Metrologo\</c>.</summary>
        public static string Racine => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Metrologo");

        /// <summary>Réglages de l'UI, configuration globale, identifiants de la BDD. Toujours
        /// en local, propre à chaque poste : on y stocke des secrets DPAPI liés au profil Windows.</summary>
        public static string Configuration => Path.Combine(Racine, "Configuration");

        /// <summary>Les sauvegardes d'imports. Surchargeable côté réseau via la clé <c>Catalogues</c>.</summary>
        public static string Catalogues =>
            AvecOverride(nameof(Catalogues), Path.Combine(Racine, "Catalogues"));

        /// <summary>
        /// Le dossier des profils d'appareils legacy (EIP / Racal / Stanford), dont les
        /// commandes GPIB peuvent s'éditer à la main. Par défaut il vit sur le réseau
        /// (<c>M:\exe_spe\Data_Metrologo\Appareils_Legacy</c>) et reste surchargeable via la
        /// clé <c>AppareilsLegacy</c>.
        /// </summary>
        public static string AppareilsLegacy =>
            AvecOverride(nameof(AppareilsLegacy), Path.Combine(BaseServeur, "Appareils_Legacy"));

        /// <summary>Les presets de stabilité. Surchargeable côté réseau via la clé <c>Presets</c>.</summary>
        public static string Presets =>
            AvecOverride(nameof(Presets), Path.Combine(Racine, "Presets"));

        /// <summary>Les caches locaux. Toujours en local : un cache n'a aucun intérêt sur le réseau.</summary>
        public static string Cache => Path.Combine(Racine, "Cache");

        /// <summary>Les modules d'incertitude. Surchargeable côté réseau via la clé <c>Incertitudes</c>.</summary>
        public static string Incertitudes =>
            AvecOverride(nameof(Incertitudes), Path.Combine(Racine, "Incertitudes"));

        /// <summary>Les archives de logs. Surchargeable côté réseau via la clé <c>ArchivesLogs</c>.</summary>
        public static string ArchivesLogs =>
            AvecOverride(nameof(ArchivesLogs), Path.Combine(Racine, "ArchivesLogs"));

        /// <summary>Les utilisateurs partagés entre postes. Surchargeable côté réseau via la clé <c>Utilisateurs</c>.</summary>
        public static string Utilisateurs =>
            AvecOverride(nameof(Utilisateurs), Path.Combine(Racine, "Utilisateurs"));

        /// <summary>Le catalogue de rubidiums partagé. Surchargeable côté réseau via la clé <c>Rubidiums</c>.</summary>
        public static string Rubidiums =>
            AvecOverride(nameof(Rubidiums), Path.Combine(Racine, "Rubidiums"));

        /// <summary>Les logs courants partagés entre postes. Surchargeable côté réseau via la clé <c>Logs</c>.</summary>
        public static string Logs =>
            AvecOverride(nameof(Logs), Path.Combine(Racine, "Logs"));

        /// <summary>Les données de Besançon (valeurs journalières et moyennes hebdo), partagées
        /// entre postes. Surchargeable côté réseau via la clé <c>Besancon</c>, qui pointe par
        /// défaut sur M:\exe_spe\Data_Metrologo\Besancon.</summary>
        public static string Besancon =>
            AvecOverride(nameof(Besancon), Path.Combine(Racine, "Besancon"));

        /// <summary>Le fichier JSON des utilisateurs, avec le hash de leur mot de passe.</summary>
        public static string FichierUtilisateurs =>
            Path.Combine(Utilisateurs, "utilisateurs.json");

        /// <summary>Le fichier JSON du catalogue de rubidiums partagé.</summary>
        public static string FichierCatalogueRubidiums =>
            Path.Combine(Rubidiums, "catalogue.json");

        /// <summary>Le fichier JSON du rubidium actif, partagé entre tous les postes comme
        /// référence commune. On le lit au démarrage pour que chaque poste reprenne le
        /// rubidium qui a été sélectionné.</summary>
        public static string FichierRubidiumActif =>
            Path.Combine(Rubidiums, "rubidium-actif.json");

        /// <summary>Le fichier JSON-lines des sessions de journal.</summary>
        public static string FichierJournalSessions =>
            Path.Combine(Logs, "sessions.jsonl");

        /// <summary>
        /// L'emplacement par défaut (le même sur tous les postes) où l'on duplique les
        /// rapports Excel des mesures. Il pointe sur le dossier Documents public de Windows
        /// (<c>C:\Users\Public\Documents\Metrologo_Backup</c>), que tous les utilisateurs du
        /// poste peuvent écrire sans droits admin, à la différence de la racine <c>C:\</c>.
        /// L'app le crée toute seule au démarrage s'il manque (voir App.OnStartup
        /// → AssurerDossierMesuresLocal).
        /// </summary>
        public static string MesuresLocalDefaut => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments),
            "Metrologo_Backup");

        /// <summary>
        /// L'emplacement local où l'on recopie les fichiers Excel produits par chaque mesure.
        /// Par défaut c'est <see cref="MesuresLocalDefaut"/>, identique sur tous les postes,
        /// mais un admin peut le rediriger ailleurs pour un poste donné via paths.config.json
        /// (clé <c>MesuresLocal</c>). Chaque rapport part vers
        /// <c>&lt;MesuresLocal&gt;\&lt;FI&gt;\Mesures_*.xlsx</c> juste après la sauvegarde
        /// principale. Comme ça, même si le réseau lâche avec le serveur principal, on garde
        /// une copie locale et on ne perd pas le rapport.
        /// </summary>
        public static string MesuresLocal =>
            AvecOverride(nameof(MesuresLocal), MesuresLocalDefaut);

        /// <summary>
        /// Vrai dès qu'un chemin local non vide est défini. Avec le défaut générique, ce sera
        /// toujours le cas, sauf si un admin a délibérément vidé le champ.
        /// </summary>
        public static bool MesuresLocalConfigure =>
            !string.IsNullOrWhiteSpace(MesuresLocal);

        /// <summary>
        /// Vérifie que le dossier <see cref="MesuresLocal"/> existe sur le poste et le crée
        /// au besoin. Idempotente, et en best-effort : si ça coince (permissions, disque
        /// inaccessible), on logge mais on ne lève pas d'exception ; la duplication des mesures
        /// se contentera d'échouer en silence lors d'une écriture ultérieure.
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
        /// Le témoin qui dit que le premier démarrage de l'app sur ce poste est déjà passé
        /// (message d'accueil affiché, raccourci bureau créé). On le range dans Configuration\
        /// pour qu'il reste propre au poste et ne soit pas synchronisé sur le réseau.
        /// </summary>
        public static string FichierFirstRunFlag =>
            Path.Combine(Configuration, "first_run_done.flag");

        /// <summary>Vrai uniquement au tout premier démarrage de l'app sur ce poste.</summary>
        public static bool EstPremierDemarrage() => !File.Exists(FichierFirstRunFlag);

        /// <summary>Note que le premier démarrage est passé, en créant le fichier témoin.</summary>
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
