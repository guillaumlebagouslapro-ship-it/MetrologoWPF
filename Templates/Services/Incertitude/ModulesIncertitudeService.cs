using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Metrologo.Models;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Incertitude
{
    /// <summary>
    /// Gère la collection des modules d'incertitude — un par fichier CSV dans
    /// <c>%LocalAppData%\Metrologo\Incertitudes\</c>.
    ///
    /// Format CSV (séparateur <c>;</c>, décimales <c>.</c>) :
    /// <code>
    /// Fonction;TempsDeMesure;BorneBasse;BorneHaute;IncertRelative;IncertAbsolue
    /// Freq;100;0.00999;10000.01;0.0;0.0000022
    /// Freq;100;10000.01;1000001.0;1.10E-11;0.0000022
    /// ...
    /// </code>
    /// Première ligne = en-têtes (ignorée). Lignes vides ou commençant par <c>#</c> ignorées.
    /// Notation scientifique <c>1.10E-11</c> acceptée.
    ///
    /// Les noms de fonctions doivent matcher la convention de l'app (cf. <c>TypeMesure</c>) :
    /// <c>Freq</c>, <c>FreqAv</c>, <c>FreqFin</c>, <c>Stab</c>, <c>Interv</c>, <c>TachyC</c>, <c>Strobo</c>.
    ///
    /// Service entièrement isolé : pour désactiver complètement la fonctionnalité, il
    /// suffit de supprimer le dossier <c>Templates/Services/Incertitude/</c> du projet.
    /// L'ancien comportement (CoeffA/B hardcodés dans <see cref="ExcelService"/>) reste
    /// inchangé tant que ce service n'est pas appelé.
    /// </summary>
    /// <summary>Résultat d'une vérification de couverture d'une valeur par un module.</summary>
    public enum CouvertureModule
    {
        /// <summary>La valeur tombe dans une ligne du module — coefficients disponibles.</summary>
        Couvert,
        /// <summary>Le module existe mais aucune ligne ne couvre la valeur (hors domaine).</summary>
        HorsPlage,
        /// <summary>Le fichier CSV du module est introuvable.</summary>
        ModuleIntrouvable
    }

    public static class ModulesIncertitudeService
    {
        public static string DossierModules => CheminsMetrologo.Incertitudes;

        // -------------------------------------------------------------------------

        /// <summary>Sous-dossier du type de mesure (sans accent ni espace, portable et lisible).</summary>
        public static string DossierPourType(TypeMesure type) => type switch
        {
            TypeMesure.Frequence       => "Frequence",
            TypeMesure.FreqAvantInterv => "FreqAvantInterv",
            TypeMesure.FreqFinale      => "FreqFinale",
            TypeMesure.Stabilite       => "Stabilite",
            TypeMesure.Interval        => "Interval",
            TypeMesure.TachyContact    => "TachymetreContact",
            TypeMesure.TachyOptique    => "TachymetreOptique",
            TypeMesure.Stroboscope     => "Stroboscope",
            _                          => string.Empty
        };

        /// <summary>
        /// Migration des anciens noms de dossier (ex: "TachyContact" → "TachymetreContact").
        /// Idempotent — appelé une seule fois au démarrage.
        /// </summary>
        public static void MigrerAnciensNomsDossiers()
        {
            var racine = DossierModules;
            if (!Directory.Exists(racine)) return;

            var migrations = new (string ancien, string nouveau)[]
            {
                ("TachyContact", "TachymetreContact"),
            };

            foreach (var (ancien, nouveau) in migrations)
            {
                string srcDir = Path.Combine(racine, ancien);
                string dstDir = Path.Combine(racine, nouveau);
                if (Directory.Exists(srcDir) && !Directory.Exists(dstDir))
                {
                    try
                    {
                        Directory.Move(srcDir, dstDir);
                        JournalLog.Info(CategorieLog.Configuration, "INCERT_DOSSIER_MIGRE",
                            $"Dossier modules renommé : {ancien} → {nouveau}");
                    }
                    catch (Exception ex)
                    {
                        JournalLog.Warn(CategorieLog.Configuration, "INCERT_DOSSIER_MIGRATION_KO",
                            $"Renommage {ancien} → {nouveau} échoué : {ex.Message}");
                    }
                }
            }
        }

        /// <summary>Chemin absolu du sous-dossier d'un type de mesure.</summary>
        public static string DossierComplet(TypeMesure type) =>
            Path.Combine(DossierModules, DossierPourType(type));

        // -------------------------------------------------------------------------

        /// <summary>
        /// Liste les modules d'un type de mesure (un sous-dossier par type). Crée le dossier
        /// si absent. Pour Frequence, dépose un CSV exemple au 1er démarrage.
        /// </summary>
        public static List<ModuleIncertitude> Lister(TypeMesure type)
        {
            string dossier = DossierComplet(type);
            if (string.IsNullOrEmpty(dossier)) return new List<ModuleIncertitude>();
            Directory.CreateDirectory(dossier);

            // Exemple uniquement dans Frequence (1ère catégorie) pour ne pas en déposer dans les 7.
            if (type == TypeMesure.Frequence
                && Directory.GetFiles(dossier, "*.csv").Length == 0)
            {
                string exemple = Path.Combine(dossier, "EXEMPLE_MF51901A.csv");
                try { File.WriteAllText(exemple, ConstruireCsvExemple(), Encoding.UTF8); }
                catch (Exception ex)
                {
                    JournalLog.Warn(CategorieLog.Configuration, "INCERT_EXEMPLE_KO",
                        $"Impossible de créer le fichier exemple : {ex.Message}");
                }
            }

            var resultats = new List<ModuleIncertitude>();
            foreach (var fichier in Directory.GetFiles(dossier, "*.csv"))
            {
                try
                {
                    var module = Charger(fichier);
                    if (module != null) resultats.Add(module);
                }
                catch (Exception ex)
                {
                    JournalLog.Warn(CategorieLog.Configuration, "INCERT_PARSE_KO",
                        $"Module {Path.GetFileName(fichier)} ignoré : {ex.Message}");
                }
            }
            return resultats.OrderBy(m => m.NumModule, StringComparer.OrdinalIgnoreCase).ToList();
        }

        /// <summary>
        /// Charge un module depuis son CSV. NumModule = nom de fichier sans extension.
        /// NomAffichage optionnel via <c># Nom: ...</c> en tête du fichier.
        /// </summary>
        public static ModuleIncertitude? Charger(string cheminCsv)
        {
            if (!File.Exists(cheminCsv)) return null;

            var module = new ModuleIncertitude
            {
                NumModule = Path.GetFileNameWithoutExtension(cheminCsv)
            };

            int numLigne = 0;
            bool headerVu = false;
            foreach (var ligne in File.ReadAllLines(cheminCsv, Encoding.UTF8))
            {
                numLigne++;
                string l = ligne.Trim();
                if (string.IsNullOrEmpty(l)) continue;
                if (l.StartsWith("#"))
                {
                    // Métadonnées : "# Nom: xxx" et "# SansTemps: true"
                    var mNom = System.Text.RegularExpressions.Regex.Match(l, @"^#\s*Nom\s*:\s*(.+)$",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (mNom.Success) module.NomAffichage = mNom.Groups[1].Value.Trim();

                    var mSansTemps = System.Text.RegularExpressions.Regex.Match(l, @"^#\s*SansTemps\s*:\s*(true|1|oui)\s*$",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (mSansTemps.Success) module.UtiliseTempsDeMesure = false;
                    continue;
                }

                // 1ère ligne non-commentaire = header (skipped)
                if (!headerVu)
                {
                    headerVu = true;
                    continue;
                }

                var champs = l.Split(';');
                // 6 col = ancien format. 9 col = nouveau format (Condition2 + Domaine2 : BB2, BH2).
                if (champs.Length < 6)
                {
                    JournalLog.Warn(CategorieLog.Configuration, "INCERT_LIGNE_KO",
                        $"{Path.GetFileName(cheminCsv)} ligne {numLigne} : au moins 6 colonnes attendues, {champs.Length} trouvées — ignorée.");
                    continue;
                }

                try
                {
                    var lg = new LigneIncertitude
                    {
                        Fonction       = champs[0].Trim(),
                        TempsDeMesure  = ParseDouble(champs[1])
                    };

                    if (champs.Length >= 9)
                    {
                        // Format 9 colonnes : Fonction;Temps;Cond2;BB1;BH1;BB2;BH2;IncRel;IncAbs
                        lg.Condition2          = champs[2].Trim();
                        lg.BorneBasse          = ParseDouble(champs[3]);
                        lg.BorneHaute          = ParseDouble(champs[4]);
                        lg.BorneBasseDomaine2  = ParseDoubleOuZero(champs[5]);
                        lg.BorneHauteDomaine2  = ParseDoubleOuZero(champs[6]);
                        lg.IncertRelative      = ParseDouble(champs[7]);
                        lg.IncertAbsolue       = ParseDouble(champs[8]);
                    }
                    else
                    {
                        // Ancien format 6 colonnes : Fonction;Temps;BB;BH;IncRel;IncAbs
                        lg.BorneBasse     = ParseDouble(champs[2]);
                        lg.BorneHaute     = ParseDouble(champs[3]);
                        lg.IncertRelative = ParseDouble(champs[4]);
                        lg.IncertAbsolue  = ParseDouble(champs[5]);
                    }
                    module.Lignes.Add(lg);
                }
                catch (Exception ex)
                {
                    JournalLog.Warn(CategorieLog.Configuration, "INCERT_LIGNE_PARSE_KO",
                        $"{Path.GetFileName(cheminCsv)} ligne {numLigne} : {ex.Message}");
                }
            }

            return module;
        }

        /// <summary>Sauvegarde un module en CSV (écrase l'existant). Appelé par l'UI Admin.</summary>
        public static void Sauvegarder(ModuleIncertitude module, TypeMesure type)
        {
            if (string.IsNullOrWhiteSpace(module.NumModule))
                throw new ArgumentException("NumModule requis pour la sauvegarde.");

            string dossier = DossierComplet(type);
            Directory.CreateDirectory(dossier);
            string chemin = Path.Combine(dossier, module.NumModule + ".csv");

            var sb = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(module.NomAffichage))
                sb.AppendLine("# Nom: " + module.NomAffichage);
            if (!module.UtiliseTempsDeMesure)
                sb.AppendLine("# SansTemps: true");
            // Format 9 colonnes : Fonction ; Temps ; Condition2 ; BB1 ; BH1 ; BB2 ; BH2 ; IncRel ; IncAbs
            sb.AppendLine("Fonction;TempsDeMesure;Condition2;BorneBasse1;BorneHaute1;BorneBasse2;BorneHaute2;IncertRelative;IncertAbsolue");
            foreach (var lg in module.Lignes)
            {
                sb.AppendLine(string.Join(";", new[]
                {
                    lg.Fonction,
                    lg.TempsDeMesure.ToString(CultureInfo.InvariantCulture),
                    lg.Condition2 ?? "",
                    lg.BorneBasse.ToString(CultureInfo.InvariantCulture),
                    lg.BorneHaute.ToString(CultureInfo.InvariantCulture),
                    lg.BorneBasseDomaine2.ToString(CultureInfo.InvariantCulture),
                    lg.BorneHauteDomaine2.ToString(CultureInfo.InvariantCulture),
                    lg.IncertRelative.ToString(CultureInfo.InvariantCulture),
                    lg.IncertAbsolue.ToString(CultureInfo.InvariantCulture)
                }));
            }
            File.WriteAllText(chemin, sb.ToString(), Encoding.UTF8);
            JournalLog.Info(CategorieLog.Administration, "INCERT_MODULE_SAUVE",
                $"Module {module.NumModule} sauvegardé dans {DossierPourType(type)} ({module.Lignes.Count} lignes).");
        }

        /// <summary>
        /// Copie le CSV d'un module vers une autre catégorie (ex. Frequence → FreqAvantInterv).
        /// Écrase la cible si elle existe — confirmation préalable attendue de l'appelant.
        /// </summary>
        public static void Copier(string numModule, TypeMesure source, TypeMesure cible)
        {
            if (source == cible)
                throw new ArgumentException("Catégories source et cible identiques.");

            string srcPath = Path.Combine(DossierComplet(source), numModule + ".csv");
            if (!File.Exists(srcPath))
                throw new FileNotFoundException("Module source introuvable : " + srcPath);

            string dossierCible = DossierComplet(cible);
            Directory.CreateDirectory(dossierCible);
            string ciblePath = Path.Combine(dossierCible, numModule + ".csv");
            File.Copy(srcPath, ciblePath, overwrite: true);

            JournalLog.Info(CategorieLog.Administration, "INCERT_MODULE_COPIE",
                $"Module {numModule} copié de {DossierPourType(source)} vers {DossierPourType(cible)}.");
        }

        /// <summary>
        /// Renomme le CSV du module dans le même sous-dossier. No-op si identique.
        /// Lève si <paramref name="nouveauNum"/> existe déjà — l'appelant doit confirmer.
        /// </summary>
        public static void Renommer(string ancienNum, string nouveauNum, TypeMesure type)
        {
            if (string.IsNullOrWhiteSpace(nouveauNum))
                throw new ArgumentException("Nouveau numéro de module requis.");
            if (string.Equals(ancienNum, nouveauNum, StringComparison.Ordinal))
                return;

            string dossier = DossierComplet(type);
            string src = Path.Combine(dossier, ancienNum + ".csv");
            string dst = Path.Combine(dossier, nouveauNum + ".csv");

            if (File.Exists(dst))
                throw new IOException($"Un module « {nouveauNum} » existe déjà dans cette catégorie.");
            if (File.Exists(src))
            {
                Directory.CreateDirectory(dossier);
                File.Move(src, dst);
            }

            JournalLog.Info(CategorieLog.Administration, "INCERT_MODULE_RENOMME",
                $"Module {ancienNum} renommé en {nouveauNum} ({DossierPourType(type)}).");
        }

        /// <summary>Supprime le fichier CSV d'un module dans le sous-dossier de son type.</summary>
        public static void Supprimer(string numModule, TypeMesure type)
        {
            string chemin = Path.Combine(DossierComplet(type), numModule + ".csv");
            if (File.Exists(chemin))
            {
                File.Delete(chemin);
                JournalLog.Warn(CategorieLog.Administration, "INCERT_MODULE_SUPPR",
                    $"Module {numModule} supprimé de {DossierPourType(type)}.");
            }
        }

        /// <summary>
        /// Retourne (CoeffA, CoeffB) pour une combinaison module/type/fonction/temps/freq.
        /// Retourne (0, 0) si module introuvable ou hors plage — le caller décide de la suite.
        /// </summary>
        public static (double CoeffA, double CoeffB) ObtenirCoefficients(
            string numModule, TypeMesure type, string fonction, double tempsSec, double freqHz)
        {
            string chemin = Path.Combine(DossierComplet(type), numModule + ".csv");
            var module = Charger(chemin);
            if (module == null) return (0, 0);

            var ligne = module.Trouver(fonction, tempsSec, freqHz);
            if (ligne == null) return (0, 0);

            return (ligne.IncertRelative, ligne.IncertAbsolue);
        }

        /// <summary>
        /// Vérifie si la valeur (Hz ou tr/min) tombe dans une ligne du module. Contrairement à
        /// <see cref="ObtenirCoefficients"/> (retourne (0,0) dans les deux cas), distingue
        /// "hors plage" de "module introuvable" — utile pour interrompre une mesure hors domaine.
        /// </summary>
        public static CouvertureModule VerifierCouverture(
            string numModule, TypeMesure type, string fonction, double tempsSec, double freqHz)
        {
            string chemin = Path.Combine(DossierComplet(type), numModule + ".csv");
            var module = Charger(chemin);
            if (module == null) return CouvertureModule.ModuleIntrouvable;

            return module.Trouver(fonction, tempsSec, freqHz) != null
                ? CouvertureModule.Couvert
                : CouvertureModule.HorsPlage;
        }

        // -------------------------------------------------------------------------

        private static double ParseDouble(string s)
        {
            // Accepte "." ou "," comme séparateur décimal.
            string normalise = s.Trim().Replace(',', '.');
            return double.Parse(normalise, NumberStyles.Float, CultureInfo.InvariantCulture);
        }

        /// <summary>Parse permissif pour les colonnes optionnelles : retourne 0 si vide ou non parsable.</summary>
        private static double ParseDoubleOuZero(string s)
        {
            string normalise = s.Trim().Replace(',', '.');
            if (string.IsNullOrEmpty(normalise)) return 0;
            return double.TryParse(normalise, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0;
        }

        private static string ConstruireCsvExemple()
        {
            // Basé sur les 12 premières lignes du module MF51901A — point de départ pour l'admin.
            var sb = new StringBuilder();
            sb.AppendLine("# Nom: Compteur de fréquence MF51901A");
            sb.AppendLine("# Format : Fonction;TempsDeMesure (s);BorneBasse (Hz);BorneHaute (Hz);IncertRelative;IncertAbsolue (Hz)");
            sb.AppendLine("# Renseigne autant de lignes que de combinaisons (Fonction × Temps × Plage de fréquence).");
            sb.AppendLine("# Le nom du fichier (sans .csv) = identifiant du module utilisé dans le code.");
            sb.AppendLine("Fonction;TempsDeMesure;BorneBasse;BorneHaute;IncertRelative;IncertAbsolue");
            sb.AppendLine("Freq;100;0.00999;10000.01;0.0;0.0000022");
            sb.AppendLine("Freq;100;10000.01;1000001.0;1.10E-11;0.0000022");
            sb.AppendLine("Freq;100;1000001.0;1000000100.0;1.10E-11;0.0");
            sb.AppendLine("Freq;50;0.01999;10000.01;0.0;0.0000026");
            sb.AppendLine("Freq;50;10000.01;1000001.0;1.10E-11;0.0000026");
            sb.AppendLine("Freq;50;1000001.0;1000000100.0;1.10E-11;0.0");
            sb.AppendLine("Freq;20;0.04999;10000.01;0.0;0.0000046");
            sb.AppendLine("Freq;20;10000.01;1000001.0;1.20E-11;0.0000046");
            sb.AppendLine("Freq;20;1000001.0;1000000100.0;1.20E-11;0.0");
            sb.AppendLine("Freq;10;0.09999;10000.01;0.0;0.0000085");
            sb.AppendLine("Freq;10;10000.01;1000001.0;2.20E-11;0.0000085");
            sb.AppendLine("Freq;10;1000001.0;1000000100.0;2.20E-11;0.0");
            sb.AppendLine("Freq;1;0.99999;10000.01;0.0;0.000017");
            sb.AppendLine("Freq;1;10000.01;1000001.0;1.10E-10;0.000017");
            sb.AppendLine("Freq;1;1000001.0;1000000100.0;1.10E-10;0.0");
            return sb.ToString();
        }
    }
}
