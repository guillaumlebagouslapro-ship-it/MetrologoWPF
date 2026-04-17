using Dapper;
using Metrologo.Models;
using Microsoft.Data.SqlClient;
using System.Threading.Tasks;

namespace Metrologo.Services
{
    public interface IAuthService
    {
        Task<Utilisateur?> AuthentifierAsync(string login, string password);
    }

    public class AuthService : IAuthService
    {
        // Pensez à mettre votre vraie chaîne de connexion ici
        private readonly string _connectionString = "Server=(localdb)\\mssqllocaldb;Database=MetrologoDB;Trusted_Connection=True;";

        public async Task<Utilisateur?> AuthentifierAsync(string login, string password)
        {
            using var db = new SqlConnection(_connectionString);

            // Requête SQL pour vérifier les identifiants
            string sql = @"
                SELECT UTI_ID as Id, UTI_LOGIN as Login, UTI_ROLE as Role 
                FROM T_UTILISATEURS 
                WHERE UTI_LOGIN = @login AND UTI_PASSWORD_HASH = @password";

            return await db.QueryFirstOrDefaultAsync<Utilisateur>(sql, new { login, password });
        }
    }
}