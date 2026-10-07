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
    /// Synchronisation du dossier d'une FI avec le partage réseau
    /// (<c>CheminsMetrologo.MesuresLocal</c>, typiquement <c>M:\…\Mesures\&lt;FI&gt;\</c>).
    ///
    /// Stratégie (voir <see cref="CheminsMetrologo.DossierFI"/>) :
    /// <list type="bullet">
    ///   <item>Réseau joignable : la FI travaille DIRECTEMENT sur le réseau, aucune copie locale,
    ///         rien à transférer (le logiciel extérieur lit le fichier à jour).</item>
    ///   <item>Réseau injoignable : la FI travaille dans Bureau\Metrologo\&lt;FI&gt;. À la fin de
    ///         chaque mesure on tente d'en pousser une copie sur le réseau, et la FI reste inscrite
    ///         dans <c>transferts_en_attente.json</c>.</item>
    ///   <item>Dès que le réseau revient (démarrage suivant, ou reprise de la FI) : le dossier local
    ///         est rapatrié sur le réseau, vérifié, puis SUPPRIMÉ — plus de double version.</item>
    /// </list>
    /// </summary>
    public static class TransfertReseauService
    {
        private static readonly object _sync = new();

        /// <summary>Chemin du fichier qui mémorise les FI en attente de transfert.</summary>
        public static string CheminFichierEnAttente =>
            Path.Combine(CheminsMetrologo.Configuration, "transferts_en_attente.json");

        /// <summary>
        /// Fin de mesure. FI sur le réseau : rien à faire (true). FI en local (réseau absent au
        /// départ ou perdu en cours de série) : pousse une copie sur le réseau si possible, et garde
        /// la FI en attente pour que le dossier local soit rapatrié puis supprimé dès qu'il n'est
        /// plus utilisé. Retourne <c>false</c> si le réseau n'a pas pu être mis à jour.
        /// </summary>
        public static async Task<bool> TransfererDossierFIAsync(string numFI)
        {
            if (string.IsNullOrWhiteSpace(numFI)) return false;

            return await Task.Run(() =>
            {
                if (CheminsMetrologo.FITravailleSurReseau(numFI))
                {
                    RetirerFIEnAttente(numFI);
                    return true;
                }

                string dossierLocal = CheminsMetrologo.DossierFILocal(numFI);
                AjouterFIEnAttente(numFI);

                if (!Directory.Exists(dossierLocal))
                {
                    Journal.Journal.Warn(CategorieLog.Systeme, "TRANSFERT_RESEAU_DOSSIER_INTROUVABLE",
                        $"Dossier local de la FI {numFI} introuvable : {dossierLocal}");
                    return false;
                }

                if (!CheminsMetrologo.ReseauMesuresJoignable())
                {
                    Journal.Journal.Warn(CategorieLog.Systeme, "TRANSFERT_RESEAU_INJOIGNABLE",
                        $"Réseau injoignable : la FI {numFI} reste en local ({dossierLocal}), "
                        + "rapatriement automatique dès le retour du réseau.");
                    return false;
                }

                try
                {
                    string dossierCible = Path.Combine(CheminsMetrologo.MesuresLocal,
                        CheminsMetrologo.NomDossierFI(numFI));
                    CopierDossierRecursif(dossierLocal, dossierCible, seulementSiPlusRecent: false);
                    Journal.Journal.Info(CategorieLog.Systeme, "TRANSFERT_RESEAU_OK",
                        $"Copie de la FI {numFI} poussée sur le réseau : {dossierCible} "
                        + "(dossier local supprimé au prochain démarrage).");
                    return true;
                }
                catch (Exception ex)
                {
                    Journal.Journal.Warn(CategorieLog.Systeme, "TRANSFERT_RESEAU_KO",
                        $"Transfert FI {numFI} vers le réseau échoué : {ex.Message} — "
                        + "FI gardée en attente de rapatriement.");
                    return false;
                }
            });
        }

        /// <summary>
        /// Rapatrie sur le réseau le dossier local de repli d'une FI, vérifie chaque fichier, puis
        /// supprime le dossier local. Ne remplace un fichier réseau que si la version locale est
        /// plus récente (un autre poste a pu travailler sur la FI entre-temps) ; les journaux sont
        /// fusionnés. Retourne <c>true</c> si la FI peut travailler sur le réseau (rien à rapatrier,
        /// ou rapatriement réussi), <c>false</c> en cas d'échec (le local est alors conservé).
        /// À n'appeler que si le dossier local n'est pas ouvert dans Excel (démarrage, reprise de FI).
        /// </summary>
        public static bool RapatrierDossierLocal(string numFI)
        {
            string dossierLocal = CheminsMetrologo.DossierFILocal(numFI);
            if (!Directory.Exists(dossierLocal)) { RetirerFIEnAttente(numFI); return true; }
            if (!Directory.EnumerateFileSystemEntries(dossierLocal).Any())
            {
                try { Directory.Delete(dossierLocal); } catch { }
                RetirerFIEnAttente(numFI);
                return true;
            }

            string dossierCible = Path.Combine(CheminsMetrologo.MesuresLocal, CheminsMetrologo.NomDossierFI(numFI));
            try
            {
                CopierDossierRecursif(dossierLocal, dossierCible, seulementSiPlusRecent: true);
                VerifierCopie(dossierLocal, dossierCible);
            }
            catch (Exception ex)
            {
                AjouterFIEnAttente(numFI);
                Journal.Journal.Warn(CategorieLog.Systeme, "RAPATRIEMENT_FI_KO",
                    $"Rapatriement de la FI {numFI} vers le réseau échoué : {ex.Message} — "
                    + $"la FI reste en local ({dossierLocal}).");
                return false;
            }

            try
            {
                Directory.Delete(dossierLocal, recursive: true);
                RetirerFIEnAttente(numFI);
                Journal.Journal.Info(CategorieLog.Systeme, "RAPATRIEMENT_FI_OK",
                    $"FI {numFI} rapatriée sur le réseau ({dossierCible}), dossier local supprimé.");
            }
            catch (Exception ex)
            {
                // Copie réussie mais un fichier local est encore ouvert : on travaille quand même
                // sur le réseau, le dossier local sera supprimé à la prochaine occasion.
                AjouterFIEnAttente(numFI);
                Journal.Journal.Warn(CategorieLog.Systeme, "RAPATRIEMENT_FI_SUPPR_KO",
                    $"FI {numFI} rapatriée mais dossier local non supprimé : {ex.Message}");
            }
            return true;
        }

        /// <summary>
        /// Au démarrage (aucun classeur ouvert) : rapatrie toutes les FI restées en local si le
        /// réseau est revenu. Best-effort : sinon elles restent en attente pour la session suivante.
        /// </summary>
        public static async Task TenterTransfertsEnAttenteAsync()
        {
            var enAttente = LireFIEnAttente();
            if (enAttente.Count == 0) return;

            if (!CheminsMetrologo.ReseauMesuresJoignable())
            {
                Journal.Journal.Info(CategorieLog.Systeme, "TRANSFERT_REPRISE_RESEAU_ABSENT",
                    $"{enAttente.Count} FI en attente de rapatriement, réseau injoignable — "
                    + "report à la prochaine session.");
                return;
            }

            Journal.Journal.Info(CategorieLog.Systeme, "TRANSFERT_REPRISE_DEBUT",
                $"Rapatriement de {enAttente.Count} FI restée(s) en local : " + string.Join(", ", enAttente));

            int nbOk = 0, nbKo = 0;
            await Task.Run(() =>
            {
                foreach (var fi in enAttente.ToList())
                {
                    if (RapatrierDossierLocal(fi)) nbOk++; else nbKo++;
                }
            });

            Journal.Journal.Info(CategorieLog.Systeme, "TRANSFERT_REPRISE_FIN",
                $"Rapatriement terminé : {nbOk} OK, {nbKo} KO (restent en attente).");
        }

        /// <summary>Inscrit la FI comme restée en local (ex. enregistrement de secours après une
        /// coupure réseau en cours de mesure).</summary>
        public static void SignalerFIEnLocal(string numFI) => AjouterFIEnAttente(numFI);

        /// <summary>Chaque fichier local doit exister côté réseau avec une taille identique (ou une
        /// version réseau plus récente, conservée volontairement). Lève sinon.</summary>
        private static void VerifierCopie(string source, string cible)
        {
            foreach (var fichier in Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories))
            {
                string relatif = Path.GetRelativePath(source, fichier);
                var src = new FileInfo(fichier);
                var dst = new FileInfo(Path.Combine(cible, relatif));
                if (!dst.Exists)
                    throw new IOException($"« {relatif} » absent sur le réseau après copie.");
                bool journal = src.Name.StartsWith("Journal_", StringComparison.OrdinalIgnoreCase);
                if (!journal && dst.LastWriteTimeUtc <= src.LastWriteTimeUtc && dst.Length != src.Length)
                    throw new IOException($"« {relatif} » : taille différente sur le réseau après copie.");
            }
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

        private static void CopierDossierRecursif(string source, string cible, bool seulementSiPlusRecent)
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
                else if (seulementSiPlusRecent && File.Exists(fichierCible)
                         && File.GetLastWriteTimeUtc(fichierCible) > File.GetLastWriteTimeUtc(fichier))
                {
                    // Rapatriement : le réseau a une version plus récente (autre poste) → on la garde.
                    Journal.Journal.Warn(CategorieLog.Systeme, "RAPATRIEMENT_RESEAU_PLUS_RECENT",
                        $"« {fichierCible} » plus récent sur le réseau que la copie locale : conservé.");
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
                CopierDossierRecursif(sousDossier, Path.Combine(cible, nomSousDossier), seulementSiPlusRecent);
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

    }
}
