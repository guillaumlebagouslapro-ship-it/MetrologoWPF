using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Résumé d'une fiche d'intervention (FI) pour le journal Admin : un accès rapide au fichier
    /// log de la FI, plus quelques infos (qui a travaillé dessus, dernière activité).
    /// </summary>
    public sealed class FicheJournalInfo
    {
        public string NumFI { get; set; } = string.Empty;
        public string CheminFichierLog { get; set; } = string.Empty;
        public string CheminDossier { get; set; } = string.Empty;
        public List<string> Operateurs { get; set; } = new();
        public DateTime DerniereActivite { get; set; }
        public int NbMesures { get; set; }

        /// <summary>Vrai si la FI n'existe que dans le dossier de repli local (Bureau\Metrologo),
        /// faute de réseau au moment de la mesure : elle sera recopiée sur le partage plus tard.</summary>
        public bool EstLocale { get; set; }

        public bool LogPresent => !string.IsNullOrEmpty(CheminFichierLog) && File.Exists(CheminFichierLog);
        public string OperateursAffiches => Operateurs.Count > 0 ? string.Join(", ", Operateurs) : "—";
        public string DerniereActiviteAffichee =>
            DerniereActivite == DateTime.MinValue ? "—" : DerniereActivite.ToString("dd/MM/yyyy HH:mm");
    }

    /// <summary>
    /// Dresse la liste des fiches d'intervention (FI) à partir des dossiers présents sous
    /// <see cref="CheminsMetrologo.MesuresLocal"/> (l'emplacement Mesures configuré, par défaut le
    /// partage réseau). Pour chaque FI, on lit son <c>Journal_&lt;FI&gt;.txt</c> pour en tirer les
    /// opérateurs et la dernière activité. Cela remplace l'ancien journal SQL de la vue Admin :
    /// chaque FI a maintenant son propre fichier log, et on se contente d'y donner un accès rapide.
    /// </summary>
    public static class JournalFIListeService
    {
        /// <summary>Dossier racine, avec un sous-dossier par FI.</summary>
        public static string DossierRacine => CheminsMetrologo.MesuresLocal;

        /// <summary>Dossier de repli des FI mesurées hors réseau (Bureau\Metrologo).</summary>
        public static string DossierRepliLocal => Path.GetDirectoryName(CheminsMetrologo.DossierFILocal("x"))!;

        /// <summary>Vrai si le dossier Mesures (réseau) est accessible au moment de l'appel.</summary>
        public static bool RacineAccessible()
        {
            try { return !string.IsNullOrWhiteSpace(DossierRacine) && Directory.Exists(DossierRacine); }
            catch { return false; }
        }

        public static List<FicheJournalInfo> Lister()
        {
            var resultats = new List<FicheJournalInfo>();
            var dejaVues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Le partage d'abord : une FI rapatriée y est la version de référence.
            if (RacineAccessible())
                AjouterFiches(DossierRacine, estLocale: false, resultats, dejaVues);

            // Puis les FI restées en local (mesurées sans réseau, pas encore transférées).
            try
            {
                if (Directory.Exists(DossierRepliLocal))
                    AjouterFiches(DossierRepliLocal, estLocale: true, resultats, dejaVues);
            }
            catch { /* Bureau illisible : on se contente du partage */ }

            return resultats.OrderByDescending(f => f.DerniereActivite).ToList();
        }

        private static void AjouterFiches(string racine, bool estLocale,
            List<FicheJournalInfo> resultats, HashSet<string> dejaVues)
        {
            IEnumerable<string> dossiers;
            try { dossiers = Directory.EnumerateDirectories(racine).ToList(); }
            catch { return; }

            foreach (var dossier in dossiers)
            {
                try
                {
                    string numFI = Path.GetFileName(dossier);
                    if (!dejaVues.Add(numFI)) continue;

                    // On attend un Journal_<FI>.txt ; à défaut, on prend le premier Journal_*.txt venu.
                    string logAttendu = Path.Combine(dossier, $"Journal_{numFI}.txt");
                    string? log = File.Exists(logAttendu)
                        ? logAttendu
                        : Directory.EnumerateFiles(dossier, "Journal_*.txt").FirstOrDefault();

                    var info = new FicheJournalInfo
                    {
                        NumFI = numFI,
                        CheminDossier = dossier,
                        CheminFichierLog = log ?? string.Empty,
                        EstLocale = estLocale,
                        DerniereActivite = log != null
                            ? File.GetLastWriteTime(log)
                            : Directory.GetLastWriteTime(dossier),
                    };

                    if (log != null) AnalyserLog(log, info);

                    resultats.Add(info);
                }
                catch { /* dossier illisible : on l'ignore */ }
            }
        }

        /// <summary>
        /// Récupère les opérateurs et le nombre de mesures réussies. Lit les deux formats :
        /// le concis (en-tête « ── 14/10/2026 · 09:00 · Jean Dupont · Baie ── », lignes « ✔ … »)
        /// et l'ancien (« Utilisateur : … », lignes MESURE_FIN).
        /// </summary>
        private static void AnalyserLog(string chemin, FicheJournalInfo info)
        {
            try
            {
                var operateurs = new List<string>();
                int nbMesures = 0;
                void AjouterOperateur(string u)
                {
                    u = u.Trim();
                    if (u.Length > 0 && u != "?" && !operateurs.Contains(u, StringComparer.OrdinalIgnoreCase))
                        operateurs.Add(u);
                }

                foreach (var ligne in File.ReadLines(chemin, Encoding.UTF8))
                {
                    if (ligne.StartsWith(JournalFIService.DebutEntete, StringComparison.Ordinal))
                    {
                        // date · heure · opérateur [· poste] ; une ligne de changement de jour n'a que la date.
                        var parties = ligne.Trim('─', ' ').Split(" · ");
                        if (parties.Length >= 3) AjouterOperateur(parties[2]);
                    }
                    else if (ligne.StartsWith("Utilisateur", StringComparison.OrdinalIgnoreCase))
                    {
                        int i = ligne.IndexOf(':');
                        if (i >= 0) AjouterOperateur(ligne[(i + 1)..]);
                    }
                    else if (ligne.Contains("MESURE_FIN", StringComparison.Ordinal)
                          || (ligne.Length > 7 && ligne[2] == ':' && ligne[7..].StartsWith(JournalFIService.MarqueOk, StringComparison.Ordinal)))
                    {
                        nbMesures++;
                    }
                }
                info.Operateurs = operateurs;
                info.NbMesures = nbMesures;
            }
            catch { /* lecture interrompue : on garde ce qu'on a déjà lu */ }
        }
    }
}
