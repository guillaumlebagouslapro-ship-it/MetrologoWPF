using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Metrologo.Services.Journal
{
    /// <summary>
    /// Résumé d'une fiche d'intervention (FI) pour le journal Admin : accès rapide au fichier
    /// log de la FI + quelques infos (opérateurs ayant travaillé dessus, dernière activité).
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
    /// Liste les fiches d'intervention (FI) à partir des dossiers présents sous
    /// <see cref="CheminsMetrologo.MesuresLocal"/> (= emplacement Mesures configuré, par défaut le
    /// partage réseau). Pour chaque FI, lit son <c>Journal_&lt;FI&gt;.txt</c> pour en extraire les
    /// opérateurs et la dernière activité. Remplace l'ancien journal SQL dans la vue Admin :
    /// chaque FI a désormais son propre fichier log, on n'offre qu'un accès rapide à ce fichier.
    /// </summary>
    public static class JournalFIListeService
    {
        /// <summary>Dossier racine contenant un sous-dossier par FI.</summary>
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

                    // Fichier journal attendu : Journal_<FI>.txt ; à défaut, tout Journal_*.txt.
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
                catch { /* dossier illisible → ignoré */ }
            }

            return resultats.OrderByDescending(f => f.DerniereActivite).ToList();
        }

        /// <summary>Extrait les opérateurs (en-têtes « Utilisateur : … ») et le nombre de mesures.</summary>
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
            catch { /* lecture partielle → on garde ce qui a été lu */ }
        }
    }
}
