using System;
using System.IO;
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
                int? maxMjd = await BesanconStore.DerniereDateJournaliereAsync(rub.Id);

                if (maxMjd == null)
                {
                    st.Niveau = NiveauSuivi.Rouge;
                    st.Titre = "Suivi Besançon — Aucune donnée";
                    st.Detail = $"Aucune valeur Besançon enregistrée pour « {rub.Designation} ».";
                    st.RapportTxt = await ConstruireRapportAsync(rub, aujourdhui, null);
                    return st;
                }

                int age = todayMjd - maxMjd.Value;
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
                        bool existe = await BesanconStore.MoyenneHebdoExisteAsync(rub.Id, mardiMjd);
                        if (!existe)
                        {
                            int nb = await BesanconStore.CompterJournalieresEntreAsync(rub.Id, debut, fin);
                            if (nb != 7) { semaineTrou = true; mardiTrou = mardiMjd; break; }
                        }
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
                st.Detail = $"Base de suivi injoignable : {ex.Message}";
                st.RapportTxt = "Impossible de lire le suivi Besançon (base SQL BASE_E2M injoignable).";
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
            sb.AppendLine("  Source           : FTP ef_utcop → base BASE_E2M (SVR-OR)");
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

                var jours = await BesanconStore.ListerJournalieresAsync(rub.Id, todayMjd - 21, todayMjd);
                sb.AppendLine("  VALEURS JOURNALIÈRES (21 derniers jours)");
                sb.AppendLine("  " + new string('-', 50));
                sb.AppendLine($"  {"MJD",-9}{"Date",-14}Valeur (Hz)");
                if (jours.Count == 0)
                    sb.AppendLine("    (aucune)");
                else
                    foreach (var j in jours)
                        sb.AppendLine($"  {j.Mjd,-9}{JourJulien.DepuisMjd(j.Mjd):dd/MM/yyyy}  "
                                    + SaisieHelper.FormaterFrequence(j.Valeur));
                sb.AppendLine();

                var moys = await BesanconStore.ListerMoyennesHebdoAsync(rub.Id, 6);
                sb.AppendLine("  MOYENNES HEBDOMADAIRES (récentes)");
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

        /// <summary>Le mardi de la semaine courante (ou aujourd'hui si l'on est un mardi).</summary>
        private static DateTime DernierMardiInclus(DateTime d)
        {
            int diff = ((int)d.DayOfWeek - (int)DayOfWeek.Tuesday + 7) % 7;
            return d.Date.AddDays(-diff);
        }
    }
}
