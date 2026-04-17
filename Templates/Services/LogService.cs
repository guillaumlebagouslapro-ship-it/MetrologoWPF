using Dapper;
using Microsoft.Data.SqlClient;
using System.Threading.Tasks;

namespace Metrologo.Services
{
    public interface ILogService
    {
        Task LogActionAsync(int? utilisateurId, string action, string details = "");
    }

    public class LogService : ILogService
    {
        private readonly string _connectionString = "Server=(localdb)\\mssqllocaldb;Database=MetrologoDB;Trusted_Connection=True;";

        public async Task LogActionAsync(int? utilisateurId, string action, string details = "")
        {
            using var db = new SqlConnection(_connectionString);
            string sql = "INSERT INTO T_LOGS (UTI_ID, LOG_ACTION, LOG_DETAILS) VALUES (@utilisateurId, @action, @details)";
            await db.ExecuteAsync(sql, new { utilisateurId, action, details });
        }
    }
}