using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Metrologo.Models;
using Metrologo.Services.Journal;

namespace Metrologo.Services
{
    /// <summary>
    /// Transfert du dossier FI complet (rapport .xlsx + Journal_FI.txt + profilings + tout
    /// le contenu de <c>C:\Users\…\Desktop\Metrologo\&lt;FI&gt;\</c>) vers le partage réseau
    /// (<c>CheminsMetrologo.MesuresLocal</c>, typiquement <c>M:\…\Mesures\&lt;FI&gt;\</c>).
    ///
    /// Stratégie :
    /// <list type="bullet">
    ///   <item>Pendant la mesure : aucune I/O réseau (tout reste en local).</item>
    ///   <item>À la fin de chaque mesure : tentative de transfert complet du dossier FI.</item>
    ///   <item>Si transfert échoue (M:\ down, latence, etc.) : on enregistre la FI dans
    ///         <c>%LocalAppData%\Metrologo\Configuration\transferts_en_attente.json</c>
    ///         pour reprise automatique au prochain démarrage de l'app.</item>
    ///   <item>Au démarrage : on rejoue tous les transferts en attente si M:\ accessible.</item>
    /// </list>
    /// </summary>
    public static class TransfertReseauService
    {
        private static readonly object _sync = new();

        /// <summary>Chemin du fichier qui mémorise les FI en attente de transfert.</summary>
        public static string CheminFichierEnAttente =>
            Path.Combine(CheminsMetrologo.Configuration, "transferts_en_attente.json");

        /// <summary>
        /// Transfère le dossier FI local (<c>Desktop\Metrologo\&lt;FI&gt;\</c>) vers le réseau.
        /// En cas d'échec ou de chemin non configuré, la FI est ajoutée à la liste en attente.
        /// Retourne <c>true</c> si succès.
        /// </summary>
        public static async Task<bool> TransfererDossierFIAsync(string numFI)
        {
            if (string.IsNullOrWhiteSpace(numFI)) return false;

            return await Task.Run(() =>
            {
                string numFISafe = SanitizerNomFichier(numFI);
                string dossierLocal = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "Metrologo", numFISafe);

                if (!Directory.Exists(dossierLocal))
                {
                    Journal.Journal.Warn(CategorieLog.Systeme, "TRANSFERT_RESEAU_DOSSIER_INTROUVABLE",
                        $"Dossier local de la FI {numFI} introuvable : {dossierLocal}");
                    return false;
                }

                if (!CheminsMetrologo.MesuresLocalConfigure)
                {
                    Journal.Journal.Warn(CategorieLog.Systeme, "TRANSFERT_RESEAU_PAS_CONFIGURE",
                        "Aucun chemin réseau configuré (Admin > Chemins d'accès) — "
                        + "le dossier reste en local.");
                    AjouterFIEnAttente(numFI);
                    return false;
                }

                try
                {
                    string dossierCible = Path.Combine(CheminsMetrologo.MesuresLocal, numFISafe);
                    Directory.CreateDirectory(dossierCible);
                    CopierDossierRecursif(dossierLocal, dossierCible);
                    RetirerFIEnAttente(numFI);
                    Journal.Journal.Info(CategorieLog.Systeme, "TRANSFERT_RESEAU_OK",
                        $"Dossier FI {numFI} transféré sur le réseau : {dossierCible}");
                    return true;
                }
                catch (Exception ex)
                {
                    AjouterFIEnAttente(numFI);
                    Journal.Journal.Warn(CategorieLog.Systeme, "TRANSFERT_RESEAU_KO",
                        $"Transfert FI {numFI} vers le réseau échoué : {ex.Message} — "
                        + "FI ajoutée à la liste de reprise au prochain démarrage.");
                    return false;
                }
            });
        }

        /// <summary>
        /// Au démarrage, rejoue tous les transferts en attente. Best-effort : si le réseau est
        /// toujours indisponible, les FI restent en liste pour la prochaine session.
        /// </summary>
        public static async Task TenterTransfertsEnAttenteAsync()
        {
            var enAttente = LireFIEnAttente();
            if (enAttente.Count == 0) return;

            if (!CheminsMetrologo.MesuresLocalConfigure)
            {
                Journal.Journal.Info(CategorieLog.Systeme, "TRANSFERT_REPRISE_PAS_CONFIGURE",
                    $"{enAttente.Count} FI en attente de transfert, mais aucun chemin réseau "
                    + "configuré — report à la prochaine session.");
                return;
            }

            Journal.Journal.Info(CategorieLog.Systeme, "TRANSFERT_REPRISE_DEBUT",
                $"Reprise auto de {enAttente.Count} transfert(s) FI en attente : "
                + string.Join(", ", enAttente));

            int nbOk = 0, nbKo = 0;
            foreach (var fi in enAttente.ToList())
            {
                bool ok = await TransfererDossierFIAsync(fi);
                if (ok) nbOk++; else nbKo++;
            }

            Journal.Journal.Info(CategorieLog.Systeme, "TRANSFERT_REPRISE_FIN",
                $"Reprise terminée : {nbOk} OK, {nbKo} KO (restent en attente).");
        }

        /// <summary>Retourne la liste des FI actuellement en attente de transfert.</summary>
        public static List<string> LireFIEnAttente()
        {
            lock (_sync)
            {
                try
                {
                    if (!File.Exists(CheminFichierEnAttente)) return new List<string>();
                    string json = File.ReadAllText(CheminFichierEnAttente);
                    var liste = JsonSerializer.Deserialize<List<string>>(json);
                    return liste ?? new List<string>();
                }
                catch (Exception ex)
                {
                    Journal.Journal.Warn(CategorieLog.Systeme, "TRANSFERT_LISTE_LECTURE_KO",
                        $"Lecture de la liste des transferts en attente échouée : {ex.Message}");
                    return new List<string>();
                }
            }
        }

        private static void AjouterFIEnAttente(string numFI)
        {
            lock (_sync)
            {
                try
                {
                    var liste = LireFIEnAttenteSansLock();
                    if (!liste.Contains(numFI, StringComparer.OrdinalIgnoreCase))
                    {
                        liste.Add(numFI);
                        SauvegarderListe(liste);
                    }
                }
                catch (Exception ex)
                {
                    Journal.Journal.Warn(CategorieLog.Systeme, "TRANSFERT_LISTE_AJOUT_KO",
                        $"Ajout FI {numFI} en attente échoué : {ex.Message}");
                }
            }
        }

        private static void RetirerFIEnAttente(string numFI)
        {
            lock (_sync)
            {
                try
                {
                    var liste = LireFIEnAttenteSansLock();
                    int avant = liste.Count;
                    liste.RemoveAll(f => string.Equals(f, numFI, StringComparison.OrdinalIgnoreCase));
                    if (liste.Count != avant) SauvegarderListe(liste);
                }
                catch { /* best-effort */ }
            }
        }

        private static List<string> LireFIEnAttenteSansLock()
        {
            if (!File.Exists(CheminFichierEnAttente)) return new List<string>();
            string json = File.ReadAllText(CheminFichierEnAttente);
            var liste = JsonSerializer.Deserialize<List<string>>(json);
            return liste ?? new List<string>();
        }

        private static void SauvegarderListe(List<string> liste)
        {
            string dossier = Path.GetDirectoryName(CheminFichierEnAttente)!;
            Directory.CreateDirectory(dossier);
            string json = JsonSerializer.Serialize(liste,
                new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(CheminFichierEnAttente, json);
        }

        private static void CopierDossierRecursif(string source, string cible)
        {
            Directory.CreateDirectory(cible);
            foreach (var fichier in Directory.EnumerateFiles(source))
            {
                string nomFichier = Path.GetFileName(fichier);
                string fichierCible = Path.Combine(cible, nomFichier);

                // Journal_<FI>.txt : FUSION obligatoire (pas d'écrasement). Plusieurs opérateurs
                // se relaient depuis des postes différents ; chaque poste n'a que ses propres blocs.
                // Un overwrite ferait perdre les sessions des autres opérateurs côté réseau.
                if (nomFichier.StartsWith("Journal_", StringComparison.OrdinalIgnoreCase)
                    && nomFichier.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                {
                    FusionnerOuCopierJournal(fichier, fichierCible);
                }
                else
                {
                    // Le local fait foi : on écrase le réseau.
                    File.Copy(fichier, fichierCible, overwrite: true);
                }
            }
            // Récursion sur les sous-dossiers (dossier FI plat aujourd'hui, prévu pour l'avenir).
            foreach (var sousDossier in Directory.EnumerateDirectories(source))
            {
                string nomSousDossier = Path.GetFileName(sousDossier);
                CopierDossierRecursif(sousDossier, Path.Combine(cible, nomSousDossier));
            }
        }

        /// <summary>
        /// Fusionne le journal FI local dans le journal réseau (simple copie si pas encore de réseau).
        /// Clé de bloc : (utilisateur, poste, date de début). Les blocs locaux remplacent leur
        /// homologue réseau (le local est plus à jour pour ses sessions) ou s'ajoutent. Résultat
        /// trié par date de début et réécrit atomiquement (temp + Move).
        /// </summary>
        private static void FusionnerOuCopierJournal(string fichierSource, string fichierCible)
        {
            // Pas encore de journal réseau → simple copie.
            if (!File.Exists(fichierCible))
            {
                File.Copy(fichierSource, fichierCible, overwrite: true);
                return;
            }

            string contenuLocal = File.ReadAllText(fichierSource, Encoding.UTF8);
            string contenuReseau = File.ReadAllText(fichierCible, Encoding.UTF8);

            var blocsReseau = DecouperBlocsSession(contenuReseau);
            var blocsLocaux = DecouperBlocsSession(contenuLocal);

            // Format inattendu (aucun bloc reconnu) → on n'écrase pas pour ne pas perdre l'historique réseau.
            if (blocsReseau.Count == 0 && blocsLocaux.Count == 0)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "JOURNAL_FUSION_FORMAT",
                    $"Fusion journal impossible (aucun bloc reconnu) pour {fichierCible} — "
                    + "journal réseau laissé intact pour ne rien perdre.");
                return;
            }

            var fusion = new Dictionary<string, BlocSession>(StringComparer.OrdinalIgnoreCase);
            void Upsert(BlocSession b, bool remplacer)
            {
                if (fusion.ContainsKey(b.Signature))
                {
                    if (remplacer) fusion[b.Signature] = b;
                }
                else
                {
                    fusion[b.Signature] = b;
                }
            }
            foreach (var b in blocsReseau) Upsert(b, remplacer: false);
            foreach (var b in blocsLocaux) Upsert(b, remplacer: true);

            // Tri par date de début ; blocs sans date parseable repoussés en fin.
            var blocsTries = fusion.Values
                .OrderBy(b => b.Debut ?? DateTime.MaxValue)
                .ToList();

            var sb = new StringBuilder();
            for (int i = 0; i < blocsTries.Count; i++)
            {
                if (i > 0) sb.Append("\r\n\r\n");
                sb.Append(blocsTries[i].Contenu.Trim('\r', '\n'));
            }
            sb.Append("\r\n");

            // Écriture atomique (temp + Move) : évite un journal réseau tronqué en cas d'interruption.
            string temp = fichierCible + ".tmp";
            File.WriteAllText(temp, sb.ToString(), Encoding.UTF8);
            File.Move(temp, fichierCible, overwrite: true);
        }

        /// <summary>
        /// Découpe un journal FI en blocs de session. Chaque bloc commence par « ===… »
        /// suivi de « Journal utilisateur — FI ».
        /// </summary>
        private static List<BlocSession> DecouperBlocsSession(string contenu)
        {
            var blocs = new List<BlocSession>();
            if (string.IsNullOrWhiteSpace(contenu)) return blocs;

            // Lookahead : découpe juste avant chaque en-tête, le séparateur reste avec son bloc.
            var separateur = new Regex(@"(?=^=+\r?\nJournal utilisateur — FI )",
                RegexOptions.Multiline);
            foreach (var morceau in separateur.Split(contenu))
            {
                if (string.IsNullOrWhiteSpace(morceau)) continue;
                if (morceau.IndexOf("Journal utilisateur — FI", StringComparison.Ordinal) < 0)
                    continue; // préambule éventuel sans en-tête reconnu → ignoré
                blocs.Add(AnalyserBloc(morceau));
            }
            return blocs;
        }

        private static BlocSession AnalyserBloc(string contenu)
        {
            string utilisateur = ExtraireChamp(contenu, "Utilisateur");
            string poste = ExtraireChamp(contenu, "Poste");
            string debutStr = ExtraireChamp(contenu, "Début session");

            DateTime? debut = null;
            if (DateTime.TryParseExact(debutStr, "dd/MM/yyyy HH:mm:ss",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out var d))
                debut = d;

            return new BlocSession
            {
                Contenu = contenu,
                Signature = $"{utilisateur}|{poste}|{debutStr}",
                Debut = debut
            };
        }

        private static string ExtraireChamp(string contenu, string libelle)
        {
            var m = Regex.Match(contenu, @"^" + Regex.Escape(libelle) + @"\s*:\s*(.*)$",
                RegexOptions.Multiline);
            return m.Success ? m.Groups[1].Value.Trim() : string.Empty;
        }

        /// <summary>Un bloc de session du journal FI (en-tête + événements), pour la fusion.</summary>
        private sealed class BlocSession
        {
            public string Contenu = string.Empty;
            public string Signature = string.Empty;
            public DateTime? Debut;
        }

        private static string SanitizerNomFichier(string nom)
        {
            if (string.IsNullOrWhiteSpace(nom)) return "sans-nom";
            var invalides = new HashSet<char>(Path.GetInvalidFileNameChars());
            var sb = new System.Text.StringBuilder(nom.Length);
            foreach (var c in nom)
            {
                sb.Append(invalides.Contains(c) ? '_' : c);
            }
            string resultat = sb.ToString().Trim(' ', '.');
            return string.IsNullOrEmpty(resultat) ? "sans-nom" : resultat;
        }
    }
}
