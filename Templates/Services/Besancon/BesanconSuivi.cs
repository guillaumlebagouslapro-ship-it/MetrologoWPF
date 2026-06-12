using System;
using System.Collections.Generic;
using System.Globalization;
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
    /// Évalue l'état du suivi Besançon (voyant vert/orange/rouge) et produit le rapport texte de
    /// l'écran principal, à partir du fichier txt cumulatif (plus aucune lecture SQL).
    /// Voyant selon l'âge de la dernière valeur journalière : 2 j max = vert, 3 à 6 j = orange,
    /// 7 j et plus = rouge.
    /// Règle métier de la moyenne hebdo de référence : moyenne des 7 valeurs d'une semaine
    /// mardi->lundi entièrement écoulée. On retient la dernière semaine complète disponible ; si la
    /// semaine en cours n'est pas calculable on garde la dernière valide. Tant qu'aucune semaine
    /// complète n'existe (première fois, données éparses), une valeur provisoire est calculée avec
    /// les valeurs disponibles. Voir CalculerReferenceHebdo.
    /// </summary>
    public static class BesanconSuiviService
    {
        private const int SeuilVertMaxJours = 2;     // jusqu'à 2 j : vert
        private const int SeuilOrangeMaxJours = 6;   // 3..6 j : orange ; au-delà : rouge
        private const int NbSemainesListe = 5;       // moyennes hebdo complètes listées (compact, tient sans scroll)
        private const int NbSemainesMaxScan = 80;    // recul max (semaines) pour retrouver une semaine complète

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
                    st.Detail = $"Aucune valeur dans le fichier « {BesanconTxtStore.CheminValeurs} ». "
                              + "Lancez une récupération (Admin → Récupérer Besançon).";
                    st.RapportTxt = await ConstruireRapportAsync(aujourdhui, valeurs, null);
                    return st;
                }

                int maxMjd = valeurs.Keys.Max();
                int age = todayMjd - maxMjd;
                st.DerniereDateMjd = maxMjd;
                st.AgeJours = age;

                var reference = CalculerReferenceHebdo(valeurs, aujourdhui);

                // Voyant basé uniquement sur la fraîcheur de la dernière valeur journalière.
                // (La moyenne de référence est détaillée dans le corps du rapport, pas répétée ici.)
                if (age > SeuilOrangeMaxJours)
                {
                    st.Niveau = NiveauSuivi.Rouge;
                    st.Titre = "Suivi Besançon — Critique";
                    st.Detail = $"Aucune nouvelle valeur depuis {age} jours.";
                }
                else if (age > SeuilVertMaxJours)
                {
                    st.Niveau = NiveauSuivi.Orange;
                    st.Titre = "Suivi Besançon — Retard";
                    st.Detail = $"Dernière valeur il y a {age} jours.";
                }
                else
                {
                    st.Niveau = NiveauSuivi.Vert;
                    st.Titre = "Suivi Besançon — À jour";
                    st.Detail = $"Données à jour — dernière valeur il y a {age} j.";
                }

                st.RapportTxt = await ConstruireRapportAsync(aujourdhui, valeurs, reference);
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

        /// <summary>Écart hebdo de référence (signé) = moyenne de la dernière semaine mardi->lundi
        /// complète (7 valeurs exactement), null si aucune. Pilote la fréquence de référence du
        /// rubidium E10-Y8 : 10 MHz x (1 + écart), cf. BesanconScheduler.</summary>
        public static async Task<double?> EcartHebdoCompletAsync(DateTime aujourdhui)
        {
            var valeurs = await BesanconTxtStore.LireAsync();
            var moys = CalculerMoyennesHebdo(valeurs, aujourdhui, 1);
            return moys.Count > 0 ? moys[0].moyenne : (double?)null;
        }

        /// <summary>Formate un écart Besançon (~1e-11 à 1e-13) en notation scientifique.
        /// SaisieHelper.FormaterFrequence, calibré pour ~10 MHz, arrondirait ces valeurs à 0.</summary>
        private static string FormaterValeur(double v) => v.ToString("0.000000E+00", CultureInfo.InvariantCulture);

        /// <summary>Construit le rapport compact (référence hebdo, fraîcheur, dernières moyennes
        /// hebdo complètes), affiché sur l'écran d'accueil et archivé sur le partage. Volontairement
        /// court pour tenir sans scroll ; le détail journalier reste dans valeurs_besancon.txt.</summary>
        private static async Task<string> ConstruireRapportAsync(
            DateTime aujourdhui, SortedDictionary<int, double> valeurs, ReferenceHebdo? reference)
        {
            var sb = new StringBuilder();

            // ligne principale : la moyenne de référence qui pilote la correction
            if (reference == null)
                sb.AppendLine("Référence hebdo : indisponible (aucune valeur)");
            else if (reference.Provisoire)
                sb.AppendLine($"Référence hebdo : {FormaterValeur(reference.Moyenne)}   (provisoire)");
            else
                sb.AppendLine($"Référence hebdo : {FormaterValeur(reference.Moyenne)}   "
                            + $"(semaine {JourJulien.DepuisMjd(reference.MardiMjd - 7):dd/MM} → {JourJulien.DepuisMjd(reference.MardiMjd - 1):dd/MM})");

            // Fraîcheur de la dernière valeur journalière.
            if (valeurs.Count > 0)
            {
                int maxMjd = valeurs.Keys.Max();
                int age = JourJulien.VersMjd(aujourdhui) - maxMjd;
                sb.AppendLine($"Dernière valeur : {JourJulien.DepuisMjd(maxMjd):dd/MM/yyyy}   (il y a {age} j)");
            }
            sb.AppendLine();

            // dernières moyennes hebdo complètes (tendance) ; la 1re ligne est la référence
            var moys = CalculerMoyennesHebdo(valeurs, aujourdhui, NbSemainesListe);
            sb.AppendLine("Moyennes hebdomadaires (mardi→lundi)");
            if (moys.Count == 0)
                sb.AppendLine("  (aucune semaine complète pour l'instant)");
            else
                foreach (var m in moys)
                    sb.AppendLine($"  {JourJulien.DepuisMjd(m.mardiMjd - 7):dd/MM} → {JourJulien.DepuisMjd(m.mardiMjd - 1):dd/MM}   {FormaterValeur(m.moyenne)}");

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

        /// <summary>Moyenne hebdomadaire de référence calculée à la volée (voir le résumé de la classe).</summary>
        private sealed class ReferenceHebdo
        {
            /// <summary>Mardi de DÉBUT de la semaine suivante ; la semaine calculée = [MardiMjd-7 ; MardiMjd-1].</summary>
            public int MardiMjd;
            public double Moyenne;
            /// <summary>True = valeur d'initialisation (pas une semaine mardi→lundi complète de 7 valeurs).</summary>
            public bool Provisoire;
        }

        /// <summary>
        /// Moyenne hebdo de référence : la plus récente semaine mardi->lundi écoulée avec exactement
        /// 7 valeurs (si la semaine en cours est incomplète, on retombe sur la dernière qui l'était).
        /// À défaut, valeur provisoire = moyenne des 7 valeurs les plus récentes, remplacée dès
        /// qu'une semaine complète existe. Null uniquement s'il n'y a aucune valeur.
        /// </summary>
        private static ReferenceHebdo? CalculerReferenceHebdo(
            SortedDictionary<int, double> valeurs, DateTime aujourdhui)
        {
            if (valeurs.Count == 0) return null;
            int todayMjd = JourJulien.VersMjd(aujourdhui);

            DateTime mardi = DernierMardiInclus(aujourdhui);
            for (int k = 0; k < NbSemainesMaxScan; k++)
            {
                int mardiMjd = JourJulien.VersMjd(mardi);
                int debut = mardiMjd - 7, fin = mardiMjd - 1;
                if (fin < todayMjd)   // semaine entièrement écoulée (mardi→lundi)
                {
                    var vals = valeurs.Where(kv => kv.Key >= debut && kv.Key <= fin)
                                      .Select(kv => kv.Value).ToList();
                    if (vals.Count == 7)
                        return new ReferenceHebdo { MardiMjd = mardiMjd, Moyenne = vals.Average(), Provisoire = false };
                }
                mardi = mardi.AddDays(-7);
            }

            // Aucune semaine complète → initialisation provisoire avec les 7 valeurs les plus récentes.
            var recentes = valeurs.OrderByDescending(kv => kv.Key).Take(7).Select(kv => kv.Value).ToList();
            return new ReferenceHebdo
            {
                MardiMjd = JourJulien.VersMjd(DernierMardiInclus(aujourdhui)),
                Moyenne = recentes.Average(),
                Provisoire = true,
            };
        }

        /// <summary>Les nbVoulu dernières moyennes hebdo complètes (mardi->lundi, 7 valeurs
        /// exactement), de la plus récente à la plus ancienne. Peut être vide s'il y a des trous.</summary>
        private static List<(int mardiMjd, double moyenne)> CalculerMoyennesHebdo(
            SortedDictionary<int, double> valeurs, DateTime aujourdhui, int nbVoulu)
        {
            var liste = new List<(int, double)>();
            int todayMjd = JourJulien.VersMjd(aujourdhui);
            DateTime mardi = DernierMardiInclus(aujourdhui);
            for (int k = 0; k < NbSemainesMaxScan && liste.Count < nbVoulu; k++)
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
