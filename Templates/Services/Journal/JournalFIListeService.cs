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

        public static List<FicheJournalInfo> Lister()
        {
            var resultats = new List<FicheJournalInfo>();
            string racine = DossierRacine;
            if (string.IsNullOrWhiteSpace(racine) || !Directory.Exists(racine))
                return resultats;

            foreach (var dossier in Directory.EnumerateDirectories(racine))
            {
                try
                {
                    string numFI = Path.GetFileName(dossier);

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
                        DerniereActivite = log != null
                            ? File.GetLastWriteTime(log)
                            : Directory.GetLastWriteTime(dossier),
                    };

                    if (log != null) AnalyserLog(log, info);

                    resultats.Add(info);
                }
                catch { /* dossier illisible : on l'ignore */ }
            }

            return resultats.OrderByDescending(f => f.DerniereActivite).ToList();
        }

        /// <summary>Récupère les opérateurs (lignes d'en-tête « Utilisateur : … ») et le nombre de mesures.</summary>
        private static void AnalyserLog(string chemin, FicheJournalInfo info)
        {
            try
            {
                var operateurs = new List<string>();
                int nbMesures = 0;
                foreach (var ligne in File.ReadLines(chemin, Encoding.UTF8))
                {
                    if (ligne.StartsWith("Utilisateur", StringComparison.OrdinalIgnoreCase))
                    {
                        int i = ligne.IndexOf(':');
                        if (i >= 0)
                        {
                            string u = ligne[(i + 1)..].Trim();
                            if (u.Length > 0 && !operateurs.Contains(u, StringComparer.OrdinalIgnoreCase))
                                operateurs.Add(u);
                        }
                    }
                    else if (ligne.Contains("MESURE_FIN", StringComparison.Ordinal))
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
