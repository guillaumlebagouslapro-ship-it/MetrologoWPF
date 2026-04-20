using System;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;
using Metrologo.Models;

namespace Metrologo.Services
{
    public interface IAuthService
    {
        Task<Utilisateur?> AuthentifierAsync(string login, string password);
    }

    public class AuthService : IAuthService
    {
        public async Task<Utilisateur?> AuthentifierAsync(string login, string password)
        {
            string passwordHash = HashPassword(password);
            using var connection = new SqliteConnection(DatabaseService.ConnectionString);
            await connection.OpenAsync();

            string sql = "SELECT Id, Login, Role FROM T_UTILISATEURS WHERE Login = @login AND PasswordHash = @passwordHash LIMIT 1";
            using var cmd = new SqliteCommand(sql, connection);
            cmd.Parameters.AddWithValue("@login", login);
            cmd.Parameters.AddWithValue("@passwordHash", passwordHash);

            using var reader = await cmd.ExecuteReaderAsync();
            if (await reader.ReadAsync())
            {
                return new Utilisateur
                {
                    Id = reader.GetInt32(0),
                    Login = reader.GetString(1),
                    Role = Enum.Parse<RoleUtilisateur>(reader.GetString(2))
                };
            }
            return null;
        }

        public static string HashPassword(string password)
        {
            using var sha = SHA256.Create();
            byte[] bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(password));
            return Convert.ToHexString(bytes);
        }
    }
}