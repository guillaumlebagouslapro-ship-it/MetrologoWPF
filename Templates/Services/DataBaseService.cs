using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;
using Metrologo.Models;
using Dapper; // N'oubliez pas l'using Dapper si vous utilisez QueryFirstOrDefaultAsync !

namespace Metrologo.Services
{
    public class DatabaseService : IDatabaseService
    {
        private static readonly string DbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "metrologo.db");
        public static string ConnectionString => $"Data Source={DbPath}";

        public async Task<bool> TesterConnexionAsync()
        {
            try
            {
                using var connection = new SqliteConnection(ConnectionString);
                await connection.OpenAsync();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public async Task<Rubidium?> GetRubidiumActifAsync()
        {
            try
            {
                // CORRECTION : Utilisation de SqliteConnection au lieu de SqlConnection
                // CORRECTION : Utilisation de ConnectionString au lieu de _connectionString
                using (var connection = new SqliteConnection(ConnectionString))
                {
                    await connection.OpenAsync();

                    string sql = @"
                        SELECT 
                            RUB_ID AS Id, 
                            RUB_DESIGNATION AS Designation, 
                            RUB_FMOYENNE AS FrequenceMoyenne, 
                            RUB_AVECGPS AS AvecGPS 
                        FROM T_RUBIDIUM 
                        WHERE RUB_ACTIF = 1 
                        LIMIT 1"; // CORRECTION : TOP 1 est pour SQL Server. En SQLite, on utilise LIMIT 1.

                    var rubidium = await connection.QueryFirstOrDefaultAsync<Rubidium>(sql);
                    return rubidium;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
            // CORRECTION : Suppression du deuxième catch redondant
        }
    }
}