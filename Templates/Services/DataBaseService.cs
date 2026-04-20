using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;
using Metrologo.Models;

namespace Metrologo.Services
{
    public class DatabaseService : IDatabaseService
    {
        private static readonly string DbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "metrologo.db");
        public static string ConnectionString => $"Data Source={DbPath}";

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
            catch { return false; }
        }

        public async Task<Rubidium?> GetRubidiumActifAsync()
        {
            try
            {
                using (var connection = new SqlConnection(_connectionString))
                {
                    await connection.OpenAsync();

                    // --- ATTENTION : À ADAPTER SELON VOTRE VRAIE BASE ---
                    // Dapper associe automatiquement le résultat de la requête "AS Id", "AS Designation" 
                    // aux propriétés de votre classe C# Rubidium !
                    string sql = @"
                        SELECT TOP 1
                            RUB_ID AS Id, 
                            RUB_DESIGNATION AS Designation, 
                            RUB_FMOYENNE AS FrequenceMoyenne, 
                            RUB_AVECGPS AS AvecGPS 
                        FROM T_RUBIDIUM 
                        WHERE RUB_ACTIF = 1"; // Adaptez le nom de la table et des colonnes !

                    var rubidium = await connection.QueryFirstOrDefaultAsync<Rubidium>(sql);
                    return rubidium;
                }
            }
            catch (Exception ex)
            {
                // Si la table n'existe pas ou que le serveur n'est pas branché, on atterrit ici
                Console.WriteLine(ex.Message);
                return null;
            }
            catch { return null; }
        }
    }
}