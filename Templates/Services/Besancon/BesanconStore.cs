using Dapper;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Metrologo.Services.Besancon
{
    /// <summary>
    /// Persistance SQL du suivi Besançon dans la base <c>BASE_E2M</c> (SVR-OR) :
    ///   - <c>T_METROLOGO_DATESRUBIS</c> (DAT_ID = date julienne MJD, RUB_ACTIF = id rubidium,
    ///     DAT_VALEUR = valeur corrigée du jour) ;
    ///   - <c>TJ_METROLOGO_SUIVIRUBI</c> (RUB_ID, RUB_AVECGPS, DAT_ID = mardi MJD, SUV_ECARTF =
    ///     moyenne hebdo, SUV_DELTATPS = moyenne × 86400 s/jour).
    ///
    /// Reproduit fidèlement la logique legacy <c>GetMoyenneHebdo</c> (F_Main.pas:3810) :
    /// moyenne sur les 7 jours précédant un mardi, EXIGE exactement 7 valeurs, SUV_DELTATPS =
    /// moyenne × 86400. Accès via <see cref="MetrologoDbService"/> + Dapper.
    /// </summary>
    public static class BesanconStore
    {
        /// <summary>
        /// Insère ou met à jour la valeur journalière (clé : DAT_ID + RUB_ACTIF). Retourne true
        /// si c'est une NOUVELLE valeur, false si une valeur existait déjà (mise à jour).
        /// </summary>
        public static async Task<bool> UpsertValeurJournaliereAsync(int rubidiumId, int mjd, double valeur)
        {
            using var c = MetrologoDbService.CreerConnexion();
            await c.OpenAsync();

            int existe = await c.ExecuteScalarAsync<int>(
                "SELECT COUNT(*) FROM T_METROLOGO_DATESRUBIS WHERE DAT_ID=@mjd AND RUB_ACTIF=@rub",
                new { mjd, rub = rubidiumId });

            if (existe > 0)
            {
                await c.ExecuteAsync(
                    "UPDATE T_METROLOGO_DATESRUBIS SET DAT_VALEUR=@v WHERE DAT_ID=@mjd AND RUB_ACTIF=@rub",
                    new { v = valeur, mjd, rub = rubidiumId });
                return false;
            }

            await c.ExecuteAsync(
                "INSERT INTO T_METROLOGO_DATESRUBIS (DAT_ID, RUB_ACTIF, DAT_VALEUR) VALUES (@mjd, @rub, @v)",
                new { mjd, rub = rubidiumId, v = valeur });
            return true;
        }

        /// <summary>
        /// Calcule la moyenne hebdo pour le mardi <paramref name="mardiMjd"/> sur les 7 jours
        /// précédents [mardiMjd-7 ; mardiMjd-1]. Retourne null si la base ne contient pas
        /// EXACTEMENT 7 valeurs (comportement legacy). Sinon (moyenne, moyenne × 86400).
        /// </summary>
        public static async Task<(double moyenne, double deltaTps)?> CalculerMoyenneHebdoAsync(int rubidiumId, int mardiMjd)
        {
            using var c = MetrologoDbService.CreerConnexion();
            await c.OpenAsync();

            int debut = mardiMjd - 7, fin = mardiMjd - 1;
            int nb = await c.ExecuteScalarAsync<int>(
                "SELECT COUNT(*) FROM T_METROLOGO_DATESRUBIS WHERE RUB_ACTIF=@rub AND DAT_ID BETWEEN @d AND @f",
                new { rub = rubidiumId, d = debut, f = fin });
            if (nb != 7) return null;

            double moyenne = await c.ExecuteScalarAsync<double>(
                "SELECT AVG(DAT_VALEUR) FROM T_METROLOGO_DATESRUBIS WHERE RUB_ACTIF=@rub AND DAT_ID BETWEEN @d AND @f",
                new { rub = rubidiumId, d = debut, f = fin });

            return (moyenne, moyenne * 86400.0);
        }

        /// <summary>Insère ou met à jour la moyenne hebdo (clé : RUB_ID + DAT_ID = mardi).</summary>
        public static async Task UpsertMoyenneHebdoAsync(int rubidiumId, bool avecGps, int mardiMjd, double moyenne, double deltaTps)
        {
            using var c = MetrologoDbService.CreerConnexion();
            await c.OpenAsync();

            int existe = await c.ExecuteScalarAsync<int>(
                "SELECT COUNT(*) FROM TJ_METROLOGO_SUIVIRUBI WHERE RUB_ID=@rub AND DAT_ID=@mjd",
                new { rub = rubidiumId, mjd = mardiMjd });

            if (existe > 0)
                await c.ExecuteAsync(
                    "UPDATE TJ_METROLOGO_SUIVIRUBI SET SUV_ECARTF=@e, SUV_DELTATPS=@dt, RUB_AVECGPS=@g "
                  + "WHERE RUB_ID=@rub AND DAT_ID=@mjd",
                    new { e = moyenne, dt = deltaTps, g = avecGps, rub = rubidiumId, mjd = mardiMjd });
            else
                await c.ExecuteAsync(
                    "INSERT INTO TJ_METROLOGO_SUIVIRUBI (RUB_ID, RUB_AVECGPS, DAT_ID, SUV_ECARTF, SUV_DELTATPS) "
                  + "VALUES (@rub, @g, @mjd, @e, @dt)",
                    new { rub = rubidiumId, g = avecGps, mjd = mardiMjd, e = moyenne, dt = deltaTps });
        }

        /// <summary>Nombre de valeurs journalières stockées pour un rubidium.</summary>
        public static async Task<int> CompterJournalieresAsync(int rubidiumId)
        {
            using var c = MetrologoDbService.CreerConnexion();
            await c.OpenAsync();
            return await c.ExecuteScalarAsync<int>(
                "SELECT COUNT(*) FROM T_METROLOGO_DATESRUBIS WHERE RUB_ACTIF=@rub", new { rub = rubidiumId });
        }

        /// <summary>Dernière moyenne hebdo calculée pour un rubidium (la plus récente), ou null.</summary>
        public static async Task<(double moyenne, int mardiMjd)?> DerniereMoyenneHebdoAsync(int rubidiumId)
        {
            using var c = MetrologoDbService.CreerConnexion();
            await c.OpenAsync();
            var row = await c.QueryFirstOrDefaultAsync(
                "SELECT TOP 1 SUV_ECARTF, DAT_ID FROM TJ_METROLOGO_SUIVIRUBI WHERE RUB_ID=@rub ORDER BY DAT_ID DESC",
                new { rub = rubidiumId });
            if (row == null) return null;
            return ((double)row.SUV_ECARTF, (int)row.DAT_ID);
        }

        /// <summary>Valeur journalière stockée pour (rubidium, date julienne), ou null si absente.</summary>
        public static async Task<double?> LireValeurJournaliereAsync(int rubidiumId, int mjd)
        {
            using var c = MetrologoDbService.CreerConnexion();
            await c.OpenAsync();
            return await c.ExecuteScalarAsync<double?>(
                "SELECT DAT_VALEUR FROM T_METROLOGO_DATESRUBIS WHERE DAT_ID=@mjd AND RUB_ACTIF=@rub",
                new { mjd, rub = rubidiumId });
        }

        /// <summary>MJD de la valeur journalière la plus récente pour un rubidium, ou null si aucune.</summary>
        public static async Task<int?> DerniereDateJournaliereAsync(int rubidiumId)
        {
            using var c = MetrologoDbService.CreerConnexion();
            await c.OpenAsync();
            return await c.ExecuteScalarAsync<int?>(
                "SELECT MAX(DAT_ID) FROM T_METROLOGO_DATESRUBIS WHERE RUB_ACTIF=@rub",
                new { rub = rubidiumId });
        }

        /// <summary>Vrai si une moyenne hebdo existe déjà pour ce mardi (DAT_ID).</summary>
        public static async Task<bool> MoyenneHebdoExisteAsync(int rubidiumId, int mardiMjd)
        {
            using var c = MetrologoDbService.CreerConnexion();
            await c.OpenAsync();
            int n = await c.ExecuteScalarAsync<int>(
                "SELECT COUNT(*) FROM TJ_METROLOGO_SUIVIRUBI WHERE RUB_ID=@rub AND DAT_ID=@mjd",
                new { rub = rubidiumId, mjd = mardiMjd });
            return n > 0;
        }

        /// <summary>Nombre de valeurs journalières présentes dans [debut ; fin] (bornes incluses).</summary>
        public static async Task<int> CompterJournalieresEntreAsync(int rubidiumId, int debut, int fin)
        {
            using var c = MetrologoDbService.CreerConnexion();
            await c.OpenAsync();
            return await c.ExecuteScalarAsync<int>(
                "SELECT COUNT(*) FROM T_METROLOGO_DATESRUBIS WHERE RUB_ACTIF=@rub AND DAT_ID BETWEEN @d AND @f",
                new { rub = rubidiumId, d = debut, f = fin });
        }

        /// <summary>Valeurs journalières de [debut ; fin], triées par date croissante.</summary>
        public static async Task<List<MesureBesancon>> ListerJournalieresAsync(int rubidiumId, int debut, int fin)
        {
            using var c = MetrologoDbService.CreerConnexion();
            await c.OpenAsync();
            var rows = await c.QueryAsync(
                "SELECT DAT_ID, DAT_VALEUR FROM T_METROLOGO_DATESRUBIS "
              + "WHERE RUB_ACTIF=@rub AND DAT_ID BETWEEN @d AND @f ORDER BY DAT_ID",
                new { rub = rubidiumId, d = debut, f = fin });
            var liste = new List<MesureBesancon>();
            foreach (var r in rows)
                liste.Add(new MesureBesancon { Mjd = (int)r.DAT_ID, Valeur = (double)r.DAT_VALEUR });
            return liste;
        }

        /// <summary>Les N dernières moyennes hebdo (mardi MJD + moyenne), de la plus récente à la plus ancienne.</summary>
        public static async Task<List<(int mardiMjd, double moyenne)>> ListerMoyennesHebdoAsync(int rubidiumId, int n)
        {
            using var c = MetrologoDbService.CreerConnexion();
            await c.OpenAsync();
            var rows = await c.QueryAsync(
                "SELECT TOP (@n) DAT_ID, SUV_ECARTF FROM TJ_METROLOGO_SUIVIRUBI "
              + "WHERE RUB_ID=@rub ORDER BY DAT_ID DESC",
                new { rub = rubidiumId, n });
            var liste = new List<(int, double)>();
            foreach (var r in rows)
                liste.Add(((int)r.DAT_ID, (double)r.SUV_ECARTF));
            return liste;
        }
    }
}
