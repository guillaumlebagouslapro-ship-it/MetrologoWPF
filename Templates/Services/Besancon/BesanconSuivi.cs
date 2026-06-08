using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Metrologo.Models;
using Metrologo.Services.Journal;

namespace Metrologo.Services.Besancon
{
    /// <summary>Niveau du voyant de suivi Besançon affiché sur l'écran principal.</summary>
    public enum NiveauSuivi
    {
        /// <summary>État indéterminé (pas de rubidium actif, ou base de suivi injoignable).</summary>
        Inconnu,
        /// <summary>Vert : données à jour.</summary>
        Vert,
        /// <summary>Orange : retard de publication encore rattrapable.</summary>
        Orange,
        /// <summary>Rouge : le calcul hebdomadaire ne peut plus être assuré.</summary>
        Rouge
    }

    /// <summary>Résultat de l'évaluation de l'état du suivi Besançon.</summary>
    public sealed class BesanconStatut
    {
        public NiveauSuivi Niveau { get; set; } = NiveauSuivi.Inconnu;
        public string Titre { get; set; } = "Suivi Besançon";
        public string Detail { get; set; } = string.Empty;
        public string RapportTxt { get; set; } = string.Empty;
        public int? DerniereDateMjd { get; set; }
        public int AgeJours { get; set; }
    }

    /// <summary>
    /// Évalue l'état du suivi Besançon (voyant vert/orange/rouge) et produit un rapport texte
    /// indenté à afficher sur l'écran principal. Lit la base partagée <c>BASE_E2M</c> — donc
    /// disponible sur TOUS les postes (la récupération FTP, elle, reste sur le poste maître).
    ///
    /// <para/>Seuils « Standard » basés sur l'ancienneté (en jours) de la dernière valeur
    /// journalière : ≤ 2 j = vert ; 3–6 j = orange (retard rattrapable) ; ≥ 7 j = rouge.
    /// Rouge également si une semaine écoulée a un trou définitif (moyenne absente ET &lt; 7
    /// valeurs présentes) → la moyenne hebdomadaire de cette semaine ne pourra plus être calculée.
    /// </summary>
    public static class BesanconSuiviService
    {
        private const int SeuilVertMaxJours = 2;     // ≤ 2 j  → vert
        private const int SeuilOrangeMaxJours = 6;   // 3..6 j → orange ; ≥ 7 j → rouge
        private const int NbSemainesControle = 6;    // nombre de semaines écoulées vérifiées
        private const int NbSemainesMaxRecul = 12;   // recul max pour collecter les moyennes du rapport

        /// <summary>Fichier texte de suivi déposé sur le partage (consultable + lu par tous les postes).</summary>
        public static string CheminRapport =>
            Path.Combine(CheminsMetrologo.Besancon, "suivi_besancon.txt");

        public static async Task<BesanconStatut> EvaluerAsync(Rubidium? rub, DateTime aujourdhui)
        {
            var st = new BesanconStatut();

            if (rub == null)
            {
                st.Niveau = NiveauSuivi.Inconnu;
                st.Detail = "Aucun rubidium actif.";
                st.RapportTxt = "Aucun rubidium actif — sélectionnez un rubidium pour suivre Besançon.";
                return st;
            }

            try
            {
                int todayMjd = JourJulien.VersMjd(aujourdhui);

                // Source unique : le fichier texte cumulatif (plus aucune lecture SQL).
                var valeurs = await BesanconTxtStore.LireAsync();

                if (valeurs.Count == 0)
                {
                    st.Niveau = NiveauSuivi.Rouge;
                    st.Titre = "Suivi Besançon — Aucune donnée";
                    st.Detail = $"Aucune valeur dans le fichier « {BesanconTxtStore.CheminValeurs} ».";
                    st.RapportTxt = await ConstruireRapportAsync(rub, aujourdhui, null);
                    return st;
                }

                int maxMjd = valeurs.Keys.Max();
                int age = todayMjd - maxMjd;
                st.DerniereDateMjd = maxMjd;
                st.AgeJours = age;

                // Détection d'une semaine écoulée dont la moyenne ne pourra plus être assurée :
                // fenêtre entièrement passée, moyenne absente ET moins de 7 valeurs présentes.
                bool semaineTrou = false;
                int? mardiTrou = null;
                DateTime mardi = DernierMardiInclus(aujourdhui);
                for (int k = 0; k < NbSemainesControle; k++)
                {
                    int mardiMjd = JourJulien.VersMjd(mardi);
                    int debut = mardiMjd - 7, fin = mardiMjd - 1;
                    if (fin < todayMjd)   // semaine entièrement écoulée
                    {
                        // Sans stockage des moyennes, une semaine écoulée est définitivement
                        // perdue si elle n'a pas ses 7 valeurs (la moyenne hebdo exige 7 jours).
                        int nb = CompterEntre(valeurs, debut, fin);
                        if (nb != 7) { semaineTrou = true; mardiTrou = mardiMjd; break; }
                    }
                    mardi = mardi.AddDays(-7);
                }

                if (age > SeuilOrangeMaxJours || semaineTrou)
                {
                    st.Niveau = NiveauSuivi.Rouge;
                    st.Titre = "Suivi Besançon — Critique";
                    st.Detail = semaineTrou
                        ? $"Semaine du mardi MJD {mardiTrou} incomplète : moyenne hebdomadaire non assurable "
                          + $"(dernière valeur il y a {age} j)."
                        : $"Aucune nouvelle valeur depuis {age} jours — le calcul hebdomadaire n'est plus assuré.";
                }
                else if (age > SeuilVertMaxJours)
                {
                    st.Niveau = NiveauSuivi.Orange;
                    st.Titre = "Suivi Besançon — Retard";
                    st.Detail = $"Dernière valeur il y a {age} jours (retard de publication). "
                              + "Calcul encore possible une fois les jours rattrapés.";
                }
                else
                {
                    st.Niveau = NiveauSuivi.Vert;
                    st.Titre = "Suivi Besançon — À jour";
                    st.Detail = $"Données à jour (dernière valeur il y a {age} j).";
                }

                st.RapportTxt = await ConstruireRapportAsync(rub, aujourdhui, maxMjd);
                return st;
            }
            catch (Exception ex)
            {
                st.Niveau = NiveauSuivi.Inconnu;
                st.Titre = "Suivi Besançon — Indisponible";
                st.Detail = $"Fichier de suivi illisible : {ex.Message}";
                st.RapportTxt = "Impossible de lire le fichier de valeurs Besançon (valeurs_besancon.txt).";
                return st;
            }
        }

        /// <summary>
        /// Construit le rapport texte indenté (valeurs journalières récentes + moyennes hebdo)
        /// et l'écrit sur le partage (best-effort). Retourne le texte produit.
        /// </summary>
        public static async Task<string> ConstruireRapportAsync(Rubidium rub, DateTime aujourdhui, int? maxMjd)
        {
            var sb = new StringBuilder();
            string filet = new string('=', 64);
            sb.AppendLine(filet);
            sb.AppendLine($"  SUIVI BESANÇON — Rubidium « {rub.Designation} »");
            sb.AppendLine(filet);
            sb.AppendLine($"  Généré le        : {aujourdhui:dd/MM/yyyy}");
            sb.AppendLine("  Source           : FTP ef_utcop → fichier valeurs_besancon.txt");
            if (maxMjd.HasValue)
            {
                int age = JourJulien.VersMjd(aujourdhui) - maxMjd.Value;
                sb.AppendLine($"  Dernière valeur  : MJD {maxMjd} "
                            + $"({JourJulien.DepuisMjd(maxMjd.Value):dd/MM/yyyy}) — il y a {age} j");
            }
            else
            {
                sb.AppendLine("  Dernière valeur  : aucune");
            }
            sb.AppendLine();

            try
            {
                int todayMjd = JourJulien.VersMjd(aujourdhui);
                var valeurs = await BesanconTxtStore.LireAsync();

                sb.AppendLine("  VALEURS JOURNALIÈRES (21 derniers jours)");
                sb.AppendLine("  " + new string('-', 50));
                sb.AppendLine($"  {"MJD",-9}{"Date",-14}Valeur (Hz)");
                var jours = valeurs.Where(kv => kv.Key >= todayMjd - 21 && kv.Key <= todayMjd).ToList();
                if (jours.Count == 0)
                    sb.AppendLine("    (aucune)");
                else
                    foreach (var j in jours)
                        sb.AppendLine($"  {j.Key,-9}{JourJulien.DepuisMjd(j.Key):dd/MM/yyyy}  "
                                    + SaisieHelper.FormaterFrequence(j.Value));
                sb.AppendLine();

                var moys = CalculerMoyennesHebdo(valeurs, aujourdhui, 6);
                sb.AppendLine("  MOYENNES HEBDOMADAIRES (recalculées depuis le fichier)");
                sb.AppendLine("  " + new string('-', 50));
                sb.AppendLine($"  {"Mardi MJD",-12}{"Date",-14}Moyenne (Hz)");
                if (moys.Count == 0)
                    sb.AppendLine("    (aucune — il faut 7 jours consécutifs)");
                else
                    foreach (var m in moys)
                        sb.AppendLine($"  {m.mardiMjd,-12}{JourJulien.DepuisMjd(m.mardiMjd):dd/MM/yyyy}  "
                                    + SaisieHelper.FormaterFrequence(m.moyenne));
            }
            catch (Exception ex)
            {
                sb.AppendLine($"  (détail indisponible : {ex.Message})");
            }

            string rapport = sb.ToString();

            // Persistance best-effort sur le partage (archive + lecture par les autres postes).
            try
            {
                Directory.CreateDirectory(CheminsMetrologo.Besancon);
                await File.WriteAllTextAsync(CheminRapport, rapport, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Journal.Journal.Warn(CategorieLog.Systeme, "BESANCON_TXT_KO",
                    $"Écriture du rapport txt Besançon échouée : {ex.Message}");
            }

            return rapport;
        }

        /// <summary>Nombre de valeurs présentes dans [debut ; fin] (bornes incluses).</summary>
        private static int CompterEntre(SortedDictionary<int, double> valeurs, int debut, int fin) =>
            valeurs.Keys.Count(mjd => mjd >= debut && mjd <= fin);

        /// <summary>
        /// Recalcule à la volée les <paramref name="nbVoulu"/> dernières moyennes hebdomadaires
        /// depuis le fichier txt : moyenne des 7 jours précédant chaque mardi, EXACTEMENT 7
        /// valeurs requises (comme le legacy <c>GetMoyenneHebdo</c>). De la plus récente à la plus
        /// ancienne ; on remonte jusqu'à <see cref="NbSemainesMaxRecul"/> mardis pour ignorer les
        /// semaines incomplètes sans s'arrêter à la première.
        /// </summary>
        private static List<(int mardiMjd, double moyenne)> CalculerMoyennesHebdo(
            SortedDictionary<int, double> valeurs, DateTime aujourdhui, int nbVoulu)
        {
            var liste = new List<(int, double)>();
            int todayMjd = JourJulien.VersMjd(aujourdhui);
            DateTime mardi = DernierMardiInclus(aujourdhui);
            for (int k = 0; k < NbSemainesMaxRecul && liste.Count < nbVoulu; k++)
            {
                int mardiMjd = JourJulien.VersMjd(mardi);
                int debut = mardiMjd - 7, fin = mardiMjd - 1;
                if (fin < todayMjd)   // semaine entièrement écoulée
                {
                    var vals = valeurs.Where(kv => kv.Key >= debut && kv.Key <= fin)
                                      .Select(kv => kv.Value).ToList();
                    if (vals.Count == 7) liste.Add((mardiMjd, vals.Average()));
                }
                mardi = mardi.AddDays(-7);
            }
            return liste;
        }

        /// <summary>Le mardi de la semaine courante (ou aujourd'hui si l'on est un mardi).</summary>
        private static DateTime DernierMardiInclus(DateTime d)
        {
            int diff = ((int)d.DayOfWeek - (int)DayOfWeek.Tuesday + 7) % 7;
            return d.Date.AddDays(-diff);
        }
    }
}
