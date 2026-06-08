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
    /// Évalue l'état du suivi Besançon (voyant vert/orange/rouge) et produit un rapport texte
    /// indenté à afficher sur l'écran principal, à partir du fichier cumulatif
    /// <see cref="BesanconTxtStore"/> (plus aucune lecture SQL).
    ///
    /// <para/>Voyant basé sur l'ancienneté (en jours) de la dernière valeur journalière :
    /// ≤ 2 j = vert ; 3–6 j = orange ; ≥ 7 j = rouge.
    ///
    /// <para/>Moyenne hebdomadaire de référence (règle métier) : moyenne des 7 valeurs d'une
    /// semaine <b>mardi → lundi</b> entièrement écoulée (exactement 7 valeurs). On retient la
    /// dernière semaine complète disponible ; si la semaine en cours n'est pas calculable, on
    /// conserve cette dernière valeur valide. Si aucune semaine complète n'existe encore (première
    /// fois / données éparses), une valeur PROVISOIRE est initialisée avec les valeurs disponibles —
    /// elle sera remplacée dès qu'une semaine complète pourra être calculée. Voir
    /// <see cref="CalculerReferenceHebdo"/>.
    /// </summary>
    public static class BesanconSuiviService
    {
        private const int SeuilVertMaxJours = 2;     // ≤ 2 j  → vert
        private const int SeuilOrangeMaxJours = 6;   // 3..6 j → orange ; ≥ 7 j → rouge
        private const int NbJoursAffiches = 14;      // valeurs journalières récentes listées dans le rapport
        private const int NbSemainesListe = 6;       // moyennes hebdo complètes listées dans le rapport
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
                    st.RapportTxt = await ConstruireRapportAsync(rub, aujourdhui, valeurs, null);
                    return st;
                }

                int maxMjd = valeurs.Keys.Max();
                int age = todayMjd - maxMjd;
                st.DerniereDateMjd = maxMjd;
                st.AgeJours = age;

                var reference = CalculerReferenceHebdo(valeurs, aujourdhui);
                string refTxt = DecrireReference(reference);

                // Voyant basé uniquement sur la fraîcheur de la dernière valeur journalière.
                if (age > SeuilOrangeMaxJours)
                {
                    st.Niveau = NiveauSuivi.Rouge;
                    st.Titre = "Suivi Besançon — Critique";
                    st.Detail = $"Aucune nouvelle valeur depuis {age} jours. Moyenne hebdo de référence : {refTxt}.";
                }
                else if (age > SeuilVertMaxJours)
                {
                    st.Niveau = NiveauSuivi.Orange;
                    st.Titre = "Suivi Besançon — Retard";
                    st.Detail = $"Dernière valeur il y a {age} jours. Moyenne hebdo de référence : {refTxt}.";
                }
                else
                {
                    st.Niveau = NiveauSuivi.Vert;
                    st.Titre = "Suivi Besançon — À jour";
                    st.Detail = $"Données à jour (dernière valeur il y a {age} j). Moyenne hebdo de référence : {refTxt}.";
                }

                st.RapportTxt = await ConstruireRapportAsync(rub, aujourdhui, valeurs, reference);
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
        /// Écart hebdomadaire de référence = moyenne de la dernière semaine mardi→lundi COMPLÈTE
        /// (exactement 7 valeurs), valeur SIGNÉE. <c>null</c> si aucune semaine complète n'est encore
        /// disponible. Sert à piloter la fréquence de référence du rubidium E10-Y8 :
        /// <c>FrequenceMoyenne = 10 MHz × (1 + écart)</c> (cf. <c>BesanconScheduler</c>).
        /// </summary>
        public static async Task<double?> EcartHebdoCompletAsync(DateTime aujourdhui)
        {
            var valeurs = await BesanconTxtStore.LireAsync();
            var moys = CalculerMoyennesHebdo(valeurs, aujourdhui, 1);
            return moys.Count > 0 ? moys[0].moyenne : (double?)null;
        }

        /// <summary>Libellé court de la moyenne hebdo de référence pour le bandeau d'accueil.</summary>
        private static string DecrireReference(ReferenceHebdo? r)
        {
            if (r == null) return "indisponible";
            return r.Provisoire
                ? $"{FormaterValeur(r.Moyenne)} (provisoire)"
                : $"{FormaterValeur(r.Moyenne)} "
                + $"(semaine du {JourJulien.DepuisMjd(r.MardiMjd - 7):dd/MM} au {JourJulien.DepuisMjd(r.MardiMjd - 1):dd/MM})";
        }

        /// <summary>
        /// Formate une valeur Besançon (écart de fréquence, de l'ordre de 1e-11 à 1e-13) en
        /// notation scientifique. <c>SaisieHelper.FormaterFrequence</c> est calibré pour des
        /// fréquences ~10 MHz et arrondirait ces très petites valeurs à « 0 ».
        /// </summary>
        private static string FormaterValeur(double v) => v.ToString("0.000000E+00", CultureInfo.InvariantCulture);

        /// <summary>
        /// Construit le rapport texte indenté (moyenne de référence en tête + valeurs journalières
        /// récentes + moyennes hebdo complètes) et l'écrit sur le partage (best-effort). Reçoit les
        /// valeurs et la référence déjà calculées par <see cref="EvaluerAsync"/>. Retourne le texte.
        /// </summary>
        private static async Task<string> ConstruireRapportAsync(
            Rubidium rub, DateTime aujourdhui, SortedDictionary<int, double> valeurs, ReferenceHebdo? reference)
        {
            var sb = new StringBuilder();
            string filet = new string('=', 64);
            sb.AppendLine(filet);
            sb.AppendLine($"  SUIVI BESANÇON — Rubidium « {rub.Designation} »");
            sb.AppendLine(filet);
            sb.AppendLine($"  Généré le        : {aujourdhui:dd/MM/yyyy}");
            sb.AppendLine("  Source           : FTP ef_utcop → fichier valeurs_besancon.txt");

            if (valeurs.Count > 0)
            {
                int maxMjd = valeurs.Keys.Max();
                int age = JourJulien.VersMjd(aujourdhui) - maxMjd;
                sb.AppendLine($"  Dernière valeur  : MJD {maxMjd} "
                            + $"({JourJulien.DepuisMjd(maxMjd):dd/MM/yyyy}) — il y a {age} j");
            }
            else
            {
                sb.AppendLine("  Dernière valeur  : aucune");
            }

            // Moyenne hebdomadaire de référence (valeur « phare » qui pilote la correction).
            if (reference == null)
                sb.AppendLine("  Moyenne hebdo    : indisponible (aucune valeur)");
            else if (reference.Provisoire)
                sb.AppendLine($"  Moyenne hebdo    : {FormaterValeur(reference.Moyenne)}  "
                            + "[PROVISOIRE — initialisée, sera recalculée sur la 1ʳᵉ semaine mardi→lundi complète]");
            else
                sb.AppendLine($"  Moyenne hebdo    : {FormaterValeur(reference.Moyenne)}  "
                            + $"[semaine du mardi {JourJulien.DepuisMjd(reference.MardiMjd - 7):dd/MM} "
                            + $"au lundi {JourJulien.DepuisMjd(reference.MardiMjd - 1):dd/MM/yyyy}]");
            sb.AppendLine();

            // Valeurs journalières : les NbJoursAffiches plus récentes présentes (jamais vide si données).
            sb.AppendLine($"  VALEURS JOURNALIÈRES ({NbJoursAffiches} dernières disponibles)");
            sb.AppendLine("  " + new string('-', 50));
            sb.AppendLine($"  {"MJD",-9}{"Date",-14}Valeur (Hz)");
            var recentes = valeurs.OrderByDescending(kv => kv.Key).Take(NbJoursAffiches)
                                  .OrderBy(kv => kv.Key).ToList();
            if (recentes.Count == 0)
                sb.AppendLine("    (aucune)");
            else
                foreach (var j in recentes)
                    sb.AppendLine($"  {j.Key,-9}{JourJulien.DepuisMjd(j.Key):dd/MM/yyyy}  "
                                + FormaterValeur(j.Value));
            sb.AppendLine();

            // Moyennes hebdomadaires COMPLÈTES récentes (peut être vide si les données ont des trous).
            var moys = CalculerMoyennesHebdo(valeurs, aujourdhui, NbSemainesListe);
            sb.AppendLine("  MOYENNES HEBDOMADAIRES COMPLÈTES (mardi→lundi, 7 valeurs)");
            sb.AppendLine("  " + new string('-', 50));
            sb.AppendLine($"  {"Mardi MJD",-12}{"Date",-14}Moyenne (Hz)");
            if (moys.Count == 0)
                sb.AppendLine("    (aucune semaine complète pour l'instant — voir « Moyenne hebdo » ci-dessus)");
            else
                foreach (var m in moys)
                    sb.AppendLine($"  {m.mardiMjd,-12}{JourJulien.DepuisMjd(m.mardiMjd):dd/MM/yyyy}  "
                                + FormaterValeur(m.moyenne));

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
        /// Détermine la moyenne hebdomadaire de référence :
        ///  1. la plus récente semaine mardi→lundi entièrement écoulée avec EXACTEMENT 7 valeurs —
        ///     c'est la « dernière valeur valide » : si la semaine en cours n'est pas complète, on
        ///     retombe automatiquement sur la dernière qui l'était ;
        ///  2. à défaut (aucune semaine complète : première fois ou données trop éparses), une
        ///     valeur PROVISOIRE = moyenne des 7 valeurs les plus récentes disponibles, remplacée
        ///     dès qu'une semaine complète pourra être calculée.
        /// Retourne null uniquement s'il n'y a aucune valeur.
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

        /// <summary>
        /// Liste des <paramref name="nbVoulu"/> dernières moyennes hebdomadaires COMPLÈTES (semaine
        /// mardi→lundi avec exactement 7 valeurs), de la plus récente à la plus ancienne. Peut être
        /// vide si les données présentent des trous.
        /// </summary>
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
