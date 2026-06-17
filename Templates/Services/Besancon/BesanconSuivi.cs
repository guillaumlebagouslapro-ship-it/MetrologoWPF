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
    /// Évalue l'état du suivi Besançon (voyant vert/orange/rouge) et produit le rapport texte,
    /// à partir du fichier txt cumulatif (plus de SQL).
    /// Voyant : age <= 2 j = vert, 3-6 j = orange, >= 7 j = rouge.
    /// Référence hebdo : moyenne de la dernière semaine mardi->lundi complète (7 valeurs) ;
    /// valeur provisoire avec les dernières disponibles si aucune semaine complète.
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

                // Source unique : le fichier txt cumulatif (plus de SQL).
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

                // Voyant sur la fraîcheur uniquement ; la moyenne est dans le rapport.
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

        /// <summary>Formate un écart (~1e-11..1e-13) en notation scientifique.
        /// FormaterFrequence arrondirait ces valeurs à 0.</summary>
        private static string FormaterValeur(double v) => v.ToString("0.000000E+00", CultureInfo.InvariantCulture);

        /// <summary>Rapport compact (référence hebdo, fraîcheur, dernières moyennes) affiché
        /// sur l'accueil et archivé. Volontairement court ; détail journalier dans valeurs_besancon.txt.</summary>
        private static async Task<string> ConstruireRapportAsync(
            DateTime aujourdhui, SortedDictionary<int, double> valeurs, ReferenceHebdo? reference)
        {
            var sb = new StringBuilder();

            // Moyenne de référence (ligne principale du rapport).
            if (reference == null)
                sb.AppendLine("Référence hebdo : indisponible (aucune valeur)");
            else if (reference.Provisoire)
                sb.AppendLine($"Référence hebdo : {FormaterValeur(reference.Moyenne)}   (provisoire)");
            else
                sb.AppendLine($"Référence hebdo : {FormaterValeur(reference.Moyenne)}   "
                            + $"(semaine {JourJulien.DepuisMjd(reference.MardiMjd - 7):dd/MM} → {JourJulien.DepuisMjd(reference.MardiMjd - 1):dd/MM})");

            // Fraîcheur de la dernière valeur.
            if (valeurs.Count > 0)
            {
                int maxMjd = valeurs.Keys.Max();
                int age = JourJulien.VersMjd(aujourdhui) - maxMjd;
                sb.AppendLine($"Dernière valeur : {JourJulien.DepuisMjd(maxMjd):dd/MM/yyyy}   (il y a {age} j)");
            }
            sb.AppendLine();

            // Dernières moyennes hebdo (tendance).
            var moys = CalculerMoyennesHebdo(valeurs, aujourdhui, NbSemainesListe);
            sb.AppendLine("Moyennes hebdomadaires (mardi→lundi)");
            if (moys.Count == 0)
                sb.AppendLine("  (aucune semaine complète pour l'instant)");
            else
                foreach (var m in moys)
                    sb.AppendLine($"  {JourJulien.DepuisMjd(m.mardiMjd - 7):dd/MM} → {JourJulien.DepuisMjd(m.mardiMjd - 1):dd/MM}   {FormaterValeur(m.moyenne)}");

            string rapport = sb.ToString();

            // Persistance best-effort sur le partage.
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

        /// <summary>Dernière semaine mardi->lundi avec 7 valeurs exactes. Valeur provisoire
        /// (7 plus récentes) si aucune semaine complète. Null si aucune valeur.</summary>
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
                if (fin < todayMjd)   // semaine entièrement écoulée
                {
                    var vals = valeurs.Where(kv => kv.Key >= debut && kv.Key <= fin)
                                      .Select(kv => kv.Value).ToList();
                    if (vals.Count == 7)
                        return new ReferenceHebdo { MardiMjd = mardiMjd, Moyenne = vals.Average(), Provisoire = false };
                }
                mardi = mardi.AddDays(-7);
            }

            // Aucune semaine complète → provisoire avec les 7 valeurs les plus récentes.
            var recentes = valeurs.OrderByDescending(kv => kv.Key).Take(7).Select(kv => kv.Value).ToList();
            return new ReferenceHebdo
            {
                MardiMjd = JourJulien.VersMjd(DernierMardiInclus(aujourdhui)),
                Moyenne = recentes.Average(),
                Provisoire = true,
            };
        }

        /// <summary>Les nbVoulu dernières moyennes hebdo complètes (mardi->lundi, 7 valeurs),
        /// de la plus récente à la plus ancienne. Vide s'il y a des trous.</summary>
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
