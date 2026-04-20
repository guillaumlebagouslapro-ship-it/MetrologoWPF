using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;
using Metrologo.Models;

namespace Metrologo.Services
{
    public class DatabaseService : IDatabaseService
    {
        // Chemin de la base de données locale (metrologo.db)
        private static readonly string DbPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "metrologo.db");

        // Propriété statique utilisée par AuthService et DatabaseInitializer
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
                using var connection = new SqliteConnection(ConnectionString);
                await connection.OpenAsync();

                string sql = @"
                    SELECT Id, Designation, FrequenceMoyenne, AvecGPS
                    FROM T_RUBIDIUM
                    WHERE AvecGPS = 1
                    LIMIT 1";

                using var cmd = new SqliteCommand(sql, connection);
                using var reader = await cmd.ExecuteReaderAsync();

                if (await reader.ReadAsync())
                {
                    return new Rubidium
                    {
                        Id = reader.GetInt32(0),
                        Designation = reader.GetString(1),
                        FrequenceMoyenne = reader.GetDouble(2),
                        AvecGPS = reader.GetBoolean(3)
                    };
                }
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DB] Erreur GetRubidium : {ex.Message}");
                return null;
            }
        }
    }
}