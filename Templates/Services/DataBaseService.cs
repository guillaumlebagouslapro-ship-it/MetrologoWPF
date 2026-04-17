using System;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using Dapper; // <-- TRÈS IMPORTANT : C'est ce qui active la magie Dapper
using Metrologo.Models;

namespace Metrologo.Services
{
    public class DatabaseService : IDatabaseService
    {
        // Pensez à mettre la VRAIE adresse de votre base de données ici !
        private readonly string _connectionString = "Server=(localdb)\\mssqllocaldb;Database=UneBaseFantomeQuiNexistePas;Trusted_Connection=True;TrustServerCertificate=True;";

        public async Task<bool> TesterConnexionAsync()
        {
            // On simule une tentative de connexion d'une demi-seconde
            await Task.Delay(500);

            // On force la réponse à "Faux" (Échec) pour voir l'interface réagir !
            return false;
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
        }
    }
}