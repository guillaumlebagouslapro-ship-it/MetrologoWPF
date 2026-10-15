using System.IO;
using System.Text.Json;

namespace Metrologo.Services
{
    /// <summary>
    /// Point d'entrée unique pour tous les chemins de fichiers de l'application
    /// (local : <c>%LocalAppData%\Metrologo\</c>, partagé : racine serveur <see cref="BaseServeur"/>).
    /// Aucun service ne doit coder un chemin en dur.
    /// <para/>
    /// <see cref="MigrerAnciensFichiers"/> (appelée par <c>App.OnStartup</c>) range dans leurs
    /// sous-dossiers les fichiers historiquement à plat ; elle est idempotente.
    /// </summary>
    public static class CheminsMetrologo
    {
        // ---------- Overrides lus depuis paths.config.json ----------

        private static Dictionary<string, string> _overrides = new();

        /// <summary>
        /// Clé particulière de paths.config.json : elle pointe vers le fichier MAÎTRE sur le
        /// partage serveur. Si elle est renseignée, ses valeurs prennent le pas sur le local
        /// au démarrage : l'admin modifie les chemins une seule fois et tous les postes suivent.
        /// </summary>
        public const string CleMasterPathsUrl = "_MasterPathsUrl";

        /// <summary>
        /// Racine du partage serveur (Incertitudes, Presets, Catalogues, ArchivesLogs, etc.).
        /// </summary>
        public const string BaseServeur = @"M:\exe_spe\Data_Metrologo";

        /// <summary>
        /// Emplacement par défaut du fichier maître. Un poste jamais configuré l'adopte au
        /// démarrage s'il répond.
        /// </summary>
        public static readonly string MasterPathsUrlDefaut =
            Path.Combine(BaseServeur, "paths.config.json");

        // Chemins serveur par défaut : ils servent à amorcer le fichier master (voir AssurerStructureServeur).
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
        /// Charge les overrides de chemins : cache local <c>Configuration\paths.config.json</c>,
        /// puis fichier maître (adopté automatiquement via <see cref="MasterPathsUrlDefaut"/> sur
        /// un poste vierge) dont les valeurs remplacent les locales et rafraîchissent le cache.
        /// Maître injoignable : on garde le cache local (mode dégradé). Idempotente.
        /// </summary>
        public static void ChargerConfigChemins()
        {
            try
            {
                // 1. Cache local
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

                // 2. Pas de MasterPathsUrl mais le défaut serveur répond : on l'adopte (poste vierge).
                if (string.IsNullOrWhiteSpace(MasterPathsUrl) && File.Exists(MasterPathsUrlDefaut))
                {
                    _overrides[CleMasterPathsUrl] = MasterPathsUrlDefaut;
                    try { SauverCacheLocal(); } catch { /* best-effort */ }
                }

                // 3. Fichier maître sur le serveur
                string masterUrl = MasterPathsUrl;
                if (!string.IsNullOrWhiteSpace(masterUrl) && File.Exists(masterUrl))
                {
                    try
                    {
                        string jsonMaster = File.ReadAllText(masterUrl);
                        var dictMaster = JsonSerializer.Deserialize<Dictionary<string, string>>(jsonMaster);
                        if (dictMaster != null)
                        {
                            // Le master a le dernier mot, sauf pour _MasterPathsUrl : on garde
                            // la locale, sinon un master qui ne la contient pas casserait le mécanisme.
                            foreach (var kv in dictMaster)
                            {
                                if (!string.Equals(kv.Key, CleMasterPathsUrl, StringComparison.OrdinalIgnoreCase))
                                    _overrides[kv.Key] = kv.Value;
                            }
                            // Cache local rafraîchi pour un prochain démarrage hors-ligne.
                            try { SauverCacheLocal(); } catch { /* best-effort */ }
                        }
                    }
                    catch
                    {
                        // Master injoignable ou corrompu : on garde le cache local (mode dégradé voulu).
                    }
                }
            }
            catch
            {
                // JSON corrompu ou illisible : chemins par défaut.
                _overrides = new Dictionary<string, string>();
            }
        }

        /// <summary>
        /// Crée sur le partage serveur les sous-dossiers de données partagées et, s'il manque,
        /// le <c>paths.config.json</c> maître (rempli avec les chemins serveur par défaut).
        /// Appelée au démarrage par tous les postes ; idempotente. Renvoie <c>true</c> si le
        /// partage est accessible, <c>false</c> si M:\ ne répond pas.
        /// </summary>
        public static bool AssurerStructureServeur()
        {
            try
            {
                // Test d'accessibilité du partage : la création de la racine échoue si M:\ n'est pas monté.
                Directory.CreateDirectory(BaseServeur);

                Directory.CreateDirectory(IncertitudesServeurDefaut);
                Directory.CreateDirectory(PresetsServeurDefaut);
                Directory.CreateDirectory(CataloguesServeurDefaut);
                Directory.CreateDirectory(ArchivesLogsServeurDefaut);
                Directory.CreateDirectory(MesuresServeurDefaut);
                Directory.CreateDirectory(UtilisateursServeurDefaut);
                Directory.CreateDirectory(RubidiumsServeurDefaut);
                Directory.CreateDirectory(LogsServeurDefaut);
                Directory.CreateDirectory(BesanconServeurDefaut);

                // Maître créé seulement s'il n'existe pas (ne pas écraser les réglages d'un admin).
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
                    // Master existant : on n'ajoute que les clés manquantes (Utilisateurs, Rubidiums,
                    // Logs, Besancon), sans toucher à ce que l'admin a configuré.
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
                // Partage indisponible : l'app reste utilisable avec les chemins locaux.
                return false;
            }
        }

        /// <summary>
        /// Enregistre les overrides dans le cache local ; si <paramref name="appliquerATousLesPostes"/>
        /// et qu'un MasterPathsUrl figure dans les overrides, écrit aussi le fichier maître (les autres
        /// postes le reprennent au démarrage). Un champ vide supprime l'override (retour au défaut).
        /// </summary>
        public static void EnregistrerConfigChemins(Dictionary<string, string> overrides,
                                                    bool appliquerATousLesPostes = false)
        {
            var aGarder = new Dictionary<string, string>();
            foreach (var kv in overrides)
            {
                if (!string.IsNullOrWhiteSpace(kv.Value))
                    aGarder[kv.Key] = kv.Value.Trim();
            }
            _overrides = aGarder;

            SauverCacheLocal();

            if (appliquerATousLesPostes && _overrides.TryGetValue(CleMasterPathsUrl, out var masterUrl)
                && !string.IsNullOrWhiteSpace(masterUrl))
            {
                try
                {
                    string? dossierMaster = Path.GetDirectoryName(masterUrl);
                    if (!string.IsNullOrEmpty(dossierMaster))
                        Directory.CreateDirectory(dossierMaster);

                    // Le master ne porte pas _MasterPathsUrl : chaque poste a la sienne.
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

        /// <summary>Sauvegarde uniquement le cache local.</summary>
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

        /// <summary>Racine locale, jamais surchargée : <c>%LocalAppData%\Metrologo\</c>.</summary>
        public static string Racine => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Metrologo");

        /// <summary>Réglages UI et configuration globale. Toujours local : contient des secrets DPAPI liés au profil Windows.</summary>
        public static string Configuration => Path.Combine(Racine, "Configuration");

        /// <summary>Catalogue d'appareils et sauvegardes d'imports. Surchargeable via la clé <c>Catalogues</c>.</summary>
        public static string Catalogues =>
            AvecOverride(nameof(Catalogues), Path.Combine(Racine, "Catalogues"));

        /// <summary>
        /// Profils d'appareils legacy (EIP / Racal / Stanford), aux commandes GPIB éditables à la
        /// main. Par défaut sur le réseau (<c>Appareils_Legacy</c>), surchargeable via <c>AppareilsLegacy</c>.
        /// </summary>
        public static string AppareilsLegacy =>
            AvecOverride(nameof(AppareilsLegacy), Path.Combine(BaseServeur, "Appareils_Legacy"));

        /// <summary>Presets de stabilité. Surchargeable via la clé <c>Presets</c>.</summary>
        public static string Presets =>
            AvecOverride(nameof(Presets), Path.Combine(Racine, "Presets"));

        /// <summary>Caches locaux (jamais sur le réseau).</summary>
        public static string Cache => Path.Combine(Racine, "Cache");

        /// <summary>Modules d'incertitude. Surchargeable via la clé <c>Incertitudes</c>.</summary>
        public static string Incertitudes =>
            AvecOverride(nameof(Incertitudes), Path.Combine(Racine, "Incertitudes"));

        /// <summary>Archives de logs. Surchargeable via la clé <c>ArchivesLogs</c>.</summary>
        public static string ArchivesLogs =>
            AvecOverride(nameof(ArchivesLogs), Path.Combine(Racine, "ArchivesLogs"));

        /// <summary>Utilisateurs partagés entre postes. Surchargeable via la clé <c>Utilisateurs</c>.</summary>
        public static string Utilisateurs =>
            AvecOverride(nameof(Utilisateurs), Path.Combine(Racine, "Utilisateurs"));

        /// <summary>Catalogue de rubidiums partagé. Surchargeable via la clé <c>Rubidiums</c>.</summary>
        public static string Rubidiums =>
            AvecOverride(nameof(Rubidiums), Path.Combine(Racine, "Rubidiums"));

        /// <summary>Logs courants partagés entre postes. Surchargeable via la clé <c>Logs</c>.</summary>
        public static string Logs =>
            AvecOverride(nameof(Logs), Path.Combine(Racine, "Logs"));

        /// <summary>Données de Besançon (valeurs journalières, moyennes hebdo), partagées entre postes.
        /// Surchargeable via la clé <c>Besancon</c> (le master serveur la fixe à <see cref="BesanconServeurDefaut"/>).</summary>
        public static string Besancon =>
            AvecOverride(nameof(Besancon), Path.Combine(Racine, "Besancon"));

        /// <summary>JSON des utilisateurs (hash des mots de passe inclus).</summary>
        public static string FichierUtilisateurs =>
            Path.Combine(Utilisateurs, "utilisateurs.json");

        /// <summary>Réglages des rapports Excel communs à tous les postes (format .xls / .xlsx), à la racine des données partagées.</summary>
        public static string FichierReglagesRapports =>
            Path.Combine(Path.GetDirectoryName(Path.TrimEndingDirectorySeparator(Utilisateurs)) ?? Racine,
                "reglages-rapports.json");

        /// <summary>JSON du catalogue de rubidiums.</summary>
        public static string FichierCatalogueRubidiums =>
            Path.Combine(Rubidiums, "catalogue.json");

        /// <summary>JSON du rubidium actif, référence commune lue au démarrage par chaque poste.</summary>
        public static string FichierRubidiumActif =>
            Path.Combine(Rubidiums, "rubidium-actif.json");

        /// <summary>
        /// Emplacement par défaut des rapports Excel des mesures : Documents public de Windows
        /// (<c>C:\Users\Public\Documents\Metrologo_Backup</c>), inscriptible sans droits admin
        /// (contrairement à la racine <c>C:\</c>). Créé au démarrage (<see cref="AssurerDossierMesuresLocal"/>).
        /// </summary>
        public static string MesuresLocalDefaut => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments),
            "Metrologo_Backup");

        /// <summary>
        /// Dossier des rapports Excel des mesures (<c>&lt;MesuresLocal&gt;\&lt;FI&gt;\Mesures_*.xlsx</c>).
        /// Par défaut <see cref="MesuresLocalDefaut"/>, redirigeable via la clé <c>MesuresLocal</c>
        /// (en pratique le dossier réseau, voir <see cref="DossierFI"/>).
        /// </summary>
        public static string MesuresLocal =>
            AvecOverride(nameof(MesuresLocal), MesuresLocalDefaut);

        /// <summary>
        /// Vrai si un chemin non vide est défini (toujours, sauf champ délibérément vidé par un admin).
        /// </summary>
        public static bool MesuresLocalConfigure =>
            !string.IsNullOrWhiteSpace(MesuresLocal);

        // ---------- Dossier de travail d'une FI : réseau si joignable, sinon local ----------

        private static readonly Dictionary<string, string> _dossiersFI = new(StringComparer.OrdinalIgnoreCase);
        private static readonly object _verrouDossiersFI = new();

        /// <summary>Nom de dossier d'une FI (caractères interdits remplacés), identique partout.</summary>
        public static string NomDossierFI(string? numFI)
        {
            if (string.IsNullOrWhiteSpace(numFI)) return "sans-nom";
            var invalides = new HashSet<char>(Path.GetInvalidFileNameChars());
            var sb = new System.Text.StringBuilder(numFI.Length);
            foreach (var c in numFI) sb.Append(invalides.Contains(c) ? '_' : c);
            string resultat = sb.ToString().Trim(' ', '.');
            return string.IsNullOrEmpty(resultat) ? "sans-nom" : resultat;
        }

        /// <summary>Dossier de REPLI d'une FI quand le réseau est injoignable : Bureau\Metrologo\&lt;FI&gt;.</summary>
        public static string DossierFILocal(string numFI) => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Metrologo", NomDossierFI(numFI));

        /// <summary>
        /// Dossier du rapport Excel et du journal d'une FI, choisi une fois par FI et par session :
        /// <list type="bullet">
        ///   <item>réseau (<see cref="MesuresLocal"/>\&lt;FI&gt;) s'il est joignable, sans copie locale
        ///         (le logiciel extérieur lit directement le fichier à jour) ; un dossier local laissé
        ///         par une session hors réseau est d'abord rapatrié puis supprimé ;</item>
        ///   <item>sinon Bureau\Metrologo\&lt;FI&gt;, recopié sur le réseau dès son retour
        ///         (<see cref="TransfertReseauService"/>).</item>
        /// </list>
        /// Potentiellement lent (test réseau, rapatriement) : à appeler hors thread UI.
        /// </summary>
        public static string DossierFI(string numFI)
        {
            string nom = NomDossierFI(numFI);
            lock (_verrouDossiersFI)
            {
                OublierSiDossierMesuresChange();
                if (_dossiersFI.TryGetValue(nom, out var deja)) return deja;

                string choisi = DossierFILocal(numFI);
                if (ReseauMesuresJoignable() && TransfertReseauService.RapatrierDossierLocal(numFI))
                    choisi = Path.Combine(MesuresLocal, nom);

                try { Directory.CreateDirectory(choisi); }
                catch
                {
                    // Réseau perdu entre le test et la création : repli local.
                    choisi = DossierFILocal(numFI);
                    Directory.CreateDirectory(choisi);
                }
                _dossiersFI[nom] = choisi;
                return choisi;
            }
        }

        private static string _dossierMesuresMemorise = string.Empty;

        /// <summary>
        /// Si le dossier « Mesures » (Admin › Chemins d'accès) a changé, oublie les dossiers de FI
        /// mémorisés et le dernier test réseau. À appeler sous _verrouDossiersFI.
        /// </summary>
        private static void OublierSiDossierMesuresChange()
        {
            string actuel = MesuresLocal ?? string.Empty;
            if (string.Equals(actuel, _dossierMesuresMemorise, StringComparison.OrdinalIgnoreCase)) return;
            _dossierMesuresMemorise = actuel;
            _dossiersFI.Clear();
            lock (_verrouTestReseau) _dateDernierTestReseau = DateTime.MinValue;
        }

        /// <summary>Vrai si la FI travaille directement sur le réseau pendant cette session.</summary>
        public static bool FITravailleSurReseau(string numFI) =>
            !string.Equals(Path.GetFullPath(DossierFI(numFI)), Path.GetFullPath(DossierFILocal(numFI)),
                           StringComparison.OrdinalIgnoreCase);

        /// <summary>Bascule la FI en local pour la suite de la session (réseau perdu en cours de
        /// série). Retourne le dossier local.</summary>
        public static string BasculerFIEnLocal(string numFI)
        {
            string local = DossierFILocal(numFI);
            Directory.CreateDirectory(local);
            lock (_verrouDossiersFI) _dossiersFI[NomDossierFI(numFI)] = local;
            return local;
        }

        private static bool _dernierResultatReseau;
        private static DateTime _dateDernierTestReseau = DateTime.MinValue;
        private static readonly object _verrouTestReseau = new();

        /// <summary>
        /// Le dossier réseau des mesures répond-il ? Un lecteur M: « endormi » met souvent plusieurs
        /// secondes à répondre : délai généreux, résultat mémorisé (60 s si joignable, 15 s sinon).
        /// Voir <see cref="ReveillerReseauMesures"/>.
        /// </summary>
        public static bool ReseauMesuresJoignable(int delaiMs = 8000)
        {
            string racine = MesuresLocal;
            if (string.IsNullOrWhiteSpace(racine)) return false;

            lock (_verrouTestReseau)
            {
                var age = DateTime.UtcNow - _dateDernierTestReseau;
                if (age < TimeSpan.FromSeconds(_dernierResultatReseau ? 60 : 15))
                    return _dernierResultatReseau;

                bool ok;
                try
                {
                    var test = Task.Run(() =>
                    {
                        if (Directory.Exists(racine)) return true;
                        Directory.CreateDirectory(racine);
                        return Directory.Exists(racine);
                    });
                    ok = test.Wait(delaiMs) && test.Result;
                }
                catch { ok = false; }

                _dernierResultatReseau = ok;
                _dateDernierTestReseau = DateTime.UtcNow;
                return ok;
            }
        }

        /// <summary>Au démarrage, en arrière-plan : réveille le lecteur réseau avant le premier test.</summary>
        public static void ReveillerReseauMesures() => Task.Run(() => ReseauMesuresJoignable());

        /// <summary>
        /// Au lancement de chaque mesure (hors thread UI) : une FI passée en local revient sur le
        /// réseau dès qu'il répond (classeur local refermé, dossier rapatrié puis supprimé).
        /// </summary>
        public static async Task ReevaluerDossierFIAsync(string numFI)
        {
            if (string.IsNullOrWhiteSpace(numFI)) return;
            string nom = NomDossierFI(numFI);
            string local = DossierFILocal(numFI);
            lock (_verrouDossiersFI)
            {
                OublierSiDossierMesuresChange();
                if (!_dossiersFI.TryGetValue(nom, out var actuel)) return; // pas encore choisi : DossierFI testera
                if (!string.Equals(Path.GetFullPath(actuel), Path.GetFullPath(local), StringComparison.OrdinalIgnoreCase))
                    return; // déjà sur le réseau
            }

            if (!await Task.Run(() => ReseauMesuresJoignable())) return;

            // Un classeur local encore ouvert empêcherait le déplacement.
            var excel = ExcelInteropHost.Instance;
            if (excel.AClasseurActif && !string.IsNullOrEmpty(excel.CheminClasseurActif)
                && Path.GetFullPath(excel.CheminClasseurActif).StartsWith(Path.GetFullPath(local), StringComparison.OrdinalIgnoreCase))
            {
                try { await excel.FermerClasseurActifAsync(); } catch { return; }
            }

            if (await Task.Run(() => TransfertReseauService.RapatrierDossierLocal(numFI)))
            {
                lock (_verrouDossiersFI) _dossiersFI[nom] = Path.Combine(MesuresLocal, nom);
            }
        }

        /// <summary>
        /// Crée <see cref="MesuresLocal"/> si besoin. Idempotente, best-effort : renvoie <c>false</c>
        /// sans lever d'exception si la création échoue.
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
        /// Témoin du premier démarrage déjà effectué sur ce poste (message d'accueil, raccourci
        /// bureau). Dans Configuration\ : propre au poste, non partagé.
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
        /// Crée sur le Bureau un raccourci .lnk « (local) » vers <see cref="MesuresLocal"/> (distinct
        /// du raccourci serveur). Idempotent. Retourne son chemin, ou null en cas d'échec.
        /// </summary>
        public static string? CreerRaccourciBureauMesuresLocal()
        {
            try
            {
                string bureau = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string cheminRaccourci = Path.Combine(bureau, "Mesures Metrologo (local).lnk");

                if (File.Exists(cheminRaccourci)) return cheminRaccourci;

                // Via COM WScript.Shell (natif Windows, aucune dépendance).
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
        /// JSON du catalogue d'appareils partagé entre postes, dans <see cref="Catalogues"/>
        /// (remplace l'ancien stockage SQL Server).
        /// </summary>
        public static string FichierCatalogueAppareils =>
            Path.Combine(Catalogues, "appareils.json");

        /// <summary>
        /// JSON des profils d'appareils legacy (EIP / Racal / Stanford), créé au 1er lancement puis
        /// relu à chaque démarrage : les commandes GPIB se corrigent sans recompiler.
        /// </summary>
        public static string FichierAppareilsLegacy =>
            Path.Combine(AppareilsLegacy, "appareils-legacy.json");

        // ---------- Migration ----------

        /// <summary>
        /// Crée les sous-dossiers locaux et y déplace les fichiers historiquement à plat dans
        /// <c>%LocalAppData%\Metrologo\</c>. Idempotent, best-effort (n'empêche pas le démarrage).
        /// </summary>
        public static void MigrerAnciensFichiers()
        {
            try
            {
                Directory.CreateDirectory(Racine);
                Directory.CreateDirectory(Configuration);
                Directory.CreateDirectory(Catalogues);
                Directory.CreateDirectory(Presets);
                Directory.CreateDirectory(Cache);
                Directory.CreateDirectory(Incertitudes);

                // Ancien nom à plat → nouveau chemin.
                DeplacerSiPresent("settings.json",         FichierSettings);
                DeplacerSiPresent("Mesures_Config.txt",     FichierMesuresConfig);
                DeplacerSiPresent("db.credentials",         FichierDbCredentials);
                DeplacerSiPresent("PresetsStabilite.json",  FichierPresetsStabilite);
                DeplacerSiPresent("journal.db",             FichierJournalDbLegacy);
                DeplacerSiPresent("journal.db-shm",         FichierJournalDbLegacy + "-shm");
                DeplacerSiPresent("journal.db-wal",         FichierJournalDbLegacy + "-wal");

                // Backups d'import du catalogue ("AppareilsCatalogue.json.imported.*") → Catalogues\.
                foreach (var f in Directory.GetFiles(Racine, "AppareilsCatalogue.json*"))
                {
                    string nom = Path.GetFileName(f);
                    string dest = Path.Combine(Catalogues, nom);
                    DeplacerFichier(f, dest);
                }
            }
            catch
            {
                // Best-effort : les services retombent sur les anciens chemins (voir ResoudreCheminAvecFallback).
            }
        }

        /// <summary>
        /// Cherche un fichier à son nouveau chemin, sinon à l'ancien (à plat dans Racine), pour le
        /// cas où <see cref="MigrerAnciensFichiers"/> n'a pas encore tourné. Retourne le 1er chemin
        /// existant, sinon <paramref name="nouveauChemin"/> (chemin de création).
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
                /* best-effort silencieux (voir ResoudreCheminAvecFallback) */
            }
        }
    }
}
