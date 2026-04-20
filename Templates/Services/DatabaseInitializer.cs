using System.Threading.Tasks;
using Microsoft.Data.Sqlite;
using Metrologo.Models;

namespace Metrologo.Services
{
    public static class DatabaseInitializer
    {
        public static async Task InitialiserAsync()
        {
            using var connection = new SqliteConnection(DatabaseService.ConnectionString);
            await connection.OpenAsync();

            string createUsers = @"
                CREATE TABLE IF NOT EXISTS T_UTILISATEURS (
                    Id           INTEGER PRIMARY KEY AUTOINCREMENT,
                    Login        TEXT NOT NULL UNIQUE,
                    PasswordHash TEXT NOT NULL,
                    Role         TEXT NOT NULL DEFAULT 'Utilisateur'
                );";

            string createRubidium = @"
                CREATE TABLE IF NOT EXISTS T_RUBIDIUM (
                    Id               INTEGER PRIMARY KEY AUTOINCREMENT,
                    Designation      TEXT NOT NULL,
                    FrequenceMoyenne REAL NOT NULL DEFAULT 10.0,
                    AvecGPS          INTEGER NOT NULL DEFAULT 0
                );";

            using (var cmd = new SqliteCommand(createUsers, connection))
                await cmd.ExecuteNonQueryAsync();

            using (var cmd = new SqliteCommand(createRubidium, connection))
                await cmd.ExecuteNonQueryAsync();

            string checkAdmin = "SELECT COUNT(*) FROM T_UTILISATEURS WHERE Login = 'admin'";
            using var checkCmd = new SqliteCommand(checkAdmin, connection);
            long count = (long)(await checkCmd.ExecuteScalarAsync())!;

            if (count == 0)
            {
                string adminHash = AuthService.HashPassword("admin123");
                string insertAdmin = @"
                    INSERT INTO T_UTILISATEURS (Login, PasswordHash, Role)
                    VALUES ('admin', @hash, 'Administrateur')";

                using var insertCmd = new SqliteCommand(insertAdmin, connection);
                insertCmd.Parameters.AddWithValue("@hash", adminHash);
                await insertCmd.ExecuteNonQueryAsync();
            }
        }
    }
}