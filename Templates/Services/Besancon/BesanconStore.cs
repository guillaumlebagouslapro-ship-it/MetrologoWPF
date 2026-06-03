using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace Metrologo.Services.Besancon
{
    /// <summary>Valeur journalière corrigée (équivalent d'une ligne de T_METROLOGO_DATESRUBIS).</summary>
    public sealed class ValeurJournaliereBesancon
    {
        public int RubidiumId { get; set; }
        public int Mjd { get; set; }
        public double Valeur { get; set; }
    }

    /// <summary>Moyenne hebdomadaire calculée (équivalent d'une ligne de TJ_METROLOGO_SUIVIRUBI).</summary>
    public sealed class MoyenneHebdoBesancon
    {
        public int RubidiumId { get; set; }
        public int MardiMjd { get; set; }
        public double Moyenne { get; set; }
        public double DeltaTpsSecondes { get; set; }
        public DateTime CalculeLe { get; set; }
    }

    public sealed class DonneesBesancon
    {
        public List<ValeurJournaliereBesancon> Journalieres { get; set; } = new();
        public List<MoyenneHebdoBesancon> Hebdos { get; set; } = new();
    }

    /// <summary>
    /// Persistance JSON (sur le partage réseau, dossier <c>Rubidiums</c>) des valeurs journalières
    /// de Besançon et des moyennes hebdomadaires. Équivalent fichier des tables legacy
    /// <c>T_METROLOGO_DATESRUBIS</c> / <c>TJ_METROLOGO_SUIVIRUBI</c> (migrable vers SQL plus tard).
    /// </summary>
    public static class BesanconStore
    {
        private static readonly object _sync = new();

        public static string Chemin => Path.Combine(CheminsMetrologo.Besancon, "besancon-suivi.json");

        public static DonneesBesancon Charger()
        {
            lock (_sync)
            {
                try
                {
                    if (File.Exists(Chemin))
                        return JsonSerializer.Deserialize<DonneesBesancon>(File.ReadAllText(Chemin))
                               ?? new DonneesBesancon();
                }
                catch { /* corrompu → repart à vide */ }
                return new DonneesBesancon();
            }
        }

        public static void Sauvegarder(DonneesBesancon d)
        {
            lock (_sync)
            {
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(Chemin)!);
                    string tmp = Chemin + ".tmp";
                    File.WriteAllText(tmp,
                        JsonSerializer.Serialize(d, new JsonSerializerOptions { WriteIndented = true }));
                    File.Move(tmp, Chemin, overwrite: true);   // écriture atomique
                }
                catch { /* best-effort */ }
            }
        }

        /// <summary>
        /// Insère ou met à jour une valeur journalière (dédoublonnage par rubidium + MJD, comme
        /// la procédure stockée legacy gérait les doublons). Retourne true si c'est une nouvelle
        /// valeur, false si on a mis à jour une valeur déjà présente.
        /// </summary>
        public static bool UpsertValeurJournaliere(DonneesBesancon d, int rubidiumId, int mjd, double valeur)
        {
            var existante = d.Journalieres.FirstOrDefault(v => v.RubidiumId == rubidiumId && v.Mjd == mjd);
            if (existante != null) { existante.Valeur = valeur; return false; }
            d.Journalieres.Add(new ValeurJournaliereBesancon { RubidiumId = rubidiumId, Mjd = mjd, Valeur = valeur });
            return true;
        }

        /// <summary>
        /// Calcule la moyenne hebdo pour le mardi <paramref name="mardiMjd"/> : moyenne des valeurs
        /// des 7 jours précédents <c>[mardiMjd-7 ; mardiMjd-1]</c>. Comme le legacy, EXIGE exactement
        /// 7 valeurs présentes — sinon retourne false (calcul impossible). <paramref name="deltaTps"/>
        /// = moyenne × 86400 (s/jour).
        /// </summary>
        public static bool CalculerMoyenneHebdo(DonneesBesancon d, int rubidiumId, int mardiMjd,
            out double moyenne, out double deltaTps)
        {
            moyenne = 0; deltaTps = 0;
            int debut = mardiMjd - 7, fin = mardiMjd - 1;
            var sur7Jours = d.Journalieres
                .Where(v => v.RubidiumId == rubidiumId && v.Mjd >= debut && v.Mjd <= fin)
                .ToList();
            if (sur7Jours.Count != 7) return false;
            moyenne = sur7Jours.Average(v => v.Valeur);
            deltaTps = moyenne * 86400.0;
            return true;
        }

        public static void UpsertMoyenneHebdo(DonneesBesancon d, int rubidiumId, int mardiMjd,
            double moyenne, double deltaTps)
        {
            var ex = d.Hebdos.FirstOrDefault(h => h.RubidiumId == rubidiumId && h.MardiMjd == mardiMjd);
            if (ex != null)
            {
                ex.Moyenne = moyenne; ex.DeltaTpsSecondes = deltaTps; ex.CalculeLe = DateTime.Now;
                return;
            }
            d.Hebdos.Add(new MoyenneHebdoBesancon
            {
                RubidiumId = rubidiumId, MardiMjd = mardiMjd,
                Moyenne = moyenne, DeltaTpsSecondes = deltaTps, CalculeLe = DateTime.Now
            });
        }
    }
}
