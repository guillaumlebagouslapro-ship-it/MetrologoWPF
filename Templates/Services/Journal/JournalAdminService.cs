using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Metrologo.Services.Journal
{
    /// <summary>Une entrée du journal d'audit administrateur.</summary>
    public sealed class EntreeJournalAdmin
    {
        public DateTime Horodatage { get; set; }
        public string Machine { get; set; } = string.Empty;
        public string Utilisateur { get; set; } = string.Empty;
        public string Action { get; set; } = string.Empty;
        public string Detail { get; set; } = string.Empty;

        public string HorodatageAffiche => Horodatage.ToString("dd/MM/yyyy HH:mm:ss");
        public string ActionLisible => JournalAdminService.LibelleAction(Action);
    }

    /// <summary>
    /// Journal d'AUDIT administrateur : trace, dans un fichier texte partagé
    /// (<c>&lt;Logs&gt;\Journal_Administration.txt</c>), uniquement les ACTIONS de configuration
    /// (changement de rubidium, création/modif/suppression de module d'incertitude, ajout/suppr
    /// d'appareil au catalogue, gestion des utilisateurs, etc.). Les simples consultations
    /// (ouverture d'écrans, lecture de FI) sont volontairement exclues.
    ///
    /// Alimenté automatiquement par la façade <see cref="Journal"/> : toute action des catégories
    /// Administration/Rubidium dont le code figure dans <see cref="_actionsAudit"/> y est ajoutée.
    /// </summary>
    public static class JournalAdminService
    {
        private static readonly object _sync = new();

        /// <summary>Fichier d'audit, sur le partage Logs (vu par tous les postes admin).</summary>
        public static string Chemin => Path.Combine(CheminsMetrologo.Logs, "Journal_Administration.txt");

        private const string Sep = " | ";

        /// <summary>
        /// Codes d'action retenus comme « audit admin » (modifications). Tout le reste — en
        /// particulier les <c>OUVERTURE_*</c> (consultations) — est ignoré.
        /// </summary>
        private static readonly HashSet<string> _actionsAudit = new(StringComparer.OrdinalIgnoreCase)
        {
            // Rubidium
            "SELECTION_RUBIDIUM", "CATALOGUE_RUBIDIUMS_MAJ",
            // Rubidium — valeur de référence (écart hebdo Besançon)
            "RUBIDIUM_VALEUR_MAJ", "RUBIDIUM_HEBDO_SOUS_SEUIL",
            // Modules d'incertitude
            "INCERT_MODULE_SAUVE", "INCERT_MODULE_COPIE", "INCERT_MODULE_SUPPR", "INCERT_MODULE_RENOMME",
            // Catalogue appareils
            "CATALOGUE_MODIF", "CATALOGUE_SUPPR", "CATALOGUE_IMPORT", "CATALOGUE_IMPORT_UI",
            // Utilisateurs / comptes
            "UTILISATEUR_AJOUTE", "UTILISATEUR_SUPPRIME", "UTILISATEUR_RENOMME", "UTILISATEUR_ROLE",
            "MDP_REINITIALISE", "MDP_MODIFIE",
            // Configuration système
            "MACRO_XLA_CONFIG", "CHEMINS_SAUVE",
            // Accès admin
            "ACCES_ADMIN_OK", "ACCES_ADMIN_KO",
        };

        public static bool EstActionAudit(string action) =>
            !string.IsNullOrEmpty(action) && _actionsAudit.Contains(action);

        /// <summary>Ajoute une entrée d'audit (best-effort, thread-safe). Enregistre aussi le nom
        /// de la machine — utilisé pour ne pas notifier le poste qui a fait le changement.</summary>
        public static void Ecrire(string action, string detail, string? utilisateur)
        {
            lock (_sync)
            {
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(Chemin)!);
                    string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string machine = Environment.MachineName;
                    string user = string.IsNullOrWhiteSpace(utilisateur) ? "?" : utilisateur;
                    // Format parseable ET lisible : ts | machine | user | action | detail
                    string d = (detail ?? string.Empty).Replace("\r", " ").Replace("\n", " ");
                    string ligne = ts + Sep + machine + Sep + user + Sep + action + Sep + d;
                    File.AppendAllText(Chemin, ligne + Environment.NewLine, Encoding.UTF8);
                }
                catch { /* best-effort : ne jamais faire échouer une action admin à cause de l'audit */ }
            }
        }

        /// <summary>Parse une ligne du journal. Gère le format actuel (5 champs avec machine) et
        /// l'ancien (4 champs sans machine). Retourne null si la ligne est inexploitable.</summary>
        private static EntreeJournalAdmin? Parser(string ligne)
        {
            if (string.IsNullOrWhiteSpace(ligne)) return null;
            var p = ligne.Split(new[] { Sep }, 5, StringSplitOptions.None);
            if (p.Length < 3) return null;
            DateTime.TryParse(p[0], out var ts);

            if (p.Length >= 5)
                return new EntreeJournalAdmin
                {
                    Horodatage = ts, Machine = p[1].Trim(), Utilisateur = p[2].Trim(),
                    Action = p[3].Trim(), Detail = p[4].Trim(),
                };

            // Ancien format : ts | user | action | detail (machine inconnue).
            return new EntreeJournalAdmin
            {
                Horodatage = ts, Machine = string.Empty, Utilisateur = p[1].Trim(),
                Action = p[2].Trim(), Detail = p.Length >= 4 ? p[3].Trim() : string.Empty,
            };
        }

        /// <summary>Toutes les entrées dans l'ordre chronologique (plus ancien → plus récent).</summary>
        public static List<EntreeJournalAdmin> LireChronologique()
        {
            var liste = new List<EntreeJournalAdmin>();
            lock (_sync)
            {
                try
                {
                    if (!File.Exists(Chemin)) return liste;
                    foreach (var ligne in File.ReadLines(Chemin, Encoding.UTF8))
                    {
                        var e = Parser(ligne);
                        if (e != null) liste.Add(e);
                    }
                }
                catch { /* lecture partielle → on garde ce qu'on a */ }
            }
            return liste;
        }

        /// <summary>Lit les entrées d'audit, les plus récentes en premier (pour le viewer).</summary>
        public static List<EntreeJournalAdmin> Lire(int max = 5000)
        {
            var liste = LireChronologique();
            liste.Reverse();
            return liste.Take(max).ToList();
        }

        /// <summary>Taille actuelle du fichier (octets) — point de départ du watcher.</summary>
        public static long Position()
        {
            try { return File.Exists(Chemin) ? new FileInfo(Chemin).Length : 0; }
            catch { return 0; }
        }

        /// <summary>
        /// Lit les NOUVELLES entrées ajoutées depuis la position <paramref name="position"/>
        /// (octets). Met à jour <paramref name="position"/> à la fin. Lecture incrémentale
        /// (seek) pour ne pas relire tout le fichier à chaque scan du watcher.
        /// </summary>
        public static List<EntreeJournalAdmin> LireDepuis(ref long position)
        {
            var nouvelles = new List<EntreeJournalAdmin>();
            lock (_sync)
            {
                try
                {
                    if (!File.Exists(Chemin)) { position = 0; return nouvelles; }
                    using var fs = new FileStream(Chemin, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    if (fs.Length < position) position = 0; // fichier tronqué/recréé
                    fs.Seek(position, SeekOrigin.Begin);
                    using var sr = new StreamReader(fs, Encoding.UTF8);
                    string? ligne;
                    while ((ligne = sr.ReadLine()) != null)
                    {
                        var e = Parser(ligne);
                        if (e != null) nouvelles.Add(e);
                    }
                    position = fs.Length;
                }
                catch { /* best-effort : on réessaiera au prochain scan */ }
            }
            return nouvelles;
        }

        /// <summary>Libellé français lisible pour un code d'action (fallback = le code brut).</summary>
        public static string LibelleAction(string action) => action switch
        {
            "SELECTION_RUBIDIUM"        => "Rubidium actif changé",
            "CATALOGUE_RUBIDIUMS_MAJ"   => "Catalogue rubidiums modifié",
            "RUBIDIUM_VALEUR_MAJ"       => "Valeur du rubidium mise à jour (écart hebdo)",
            "RUBIDIUM_HEBDO_SOUS_SEUIL" => "⚠ Écart hebdo rubidium sous la limite",
            "INCERT_MODULE_SAUVE"     => "Module d'incertitude enregistré",
            "INCERT_MODULE_COPIE"     => "Module d'incertitude dupliqué",
            "INCERT_MODULE_SUPPR"     => "Module d'incertitude supprimé",
            "INCERT_MODULE_RENOMME"   => "Module d'incertitude renommé",
            "CATALOGUE_MODIF"         => "Appareil ajouté / modifié au catalogue",
            "CATALOGUE_SUPPR"         => "Appareil supprimé du catalogue",
            "CATALOGUE_IMPORT"        => "Catalogue appareils importé",
            "CATALOGUE_IMPORT_UI"     => "Catalogue appareils importé",
            "UTILISATEUR_AJOUTE"      => "Utilisateur ajouté",
            "UTILISATEUR_SUPPRIME"    => "Utilisateur supprimé",
            "UTILISATEUR_RENOMME"     => "Utilisateur renommé",
            "UTILISATEUR_ROLE"        => "Rôle utilisateur modifié",
            "MDP_REINITIALISE"        => "Mot de passe réinitialisé",
            "MDP_MODIFIE"             => "Mot de passe modifié",
            "MACRO_XLA_CONFIG"        => "Chemin macro Excel modifié",
            "CHEMINS_SAUVE"           => "Chemins de stockage modifiés",
            "ACCES_ADMIN_OK"          => "Accès administrateur accordé",
            "ACCES_ADMIN_KO"          => "Accès administrateur refusé",
            _                         => action,
        };
    }
}
