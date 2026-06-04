using Dapper;
using Metrologo.Services.Journal;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services
{
    /// <summary>
    /// Accès à la base SQL <b>Metrologo</b> (SQL Server <c>SVR-OR</c>, login <c>MetrologoUser</c>).
    /// Contient les tables historiques de suivi Besançon :
    /// <c>T_METROLOGO_DATESRUBIS</c> (valeurs journalières), <c>TJ_METROLOGO_SUIVIRUBI</c>
    /// (moyennes hebdo) et <c>TR_METROLOGO_RUBIDIUMS</c>.
    ///
    /// <para/>Pattern identique à <see cref="AseriService"/> : <see cref="SqlConnection"/> +
    /// Dapper. La chaîne est en dur (base interne au LAN, dépôt privé) — comme ASERi.
    /// </summary>
    public static class MetrologoDbService
    {
        // Serveur SVR-OR, base BASE_E2M (contient les tables T_METROLOGO_*).
        // TrustServerCertificate=true pour les SQL Server anciens sans certificat valide.
        private const string ConnectionString =
            "Server=SVR-OR;Database=BASE_E2M;User Id=MetrologoUser;Password=Metro2026;"
          + "TrustServerCertificate=true;Encrypt=false;Connect Timeout=6";

        /// <summary>Crée une connexion (à utiliser dans un <c>using</c> + Dapper).</summary>
        public static SqlConnection CreerConnexion() => new SqlConnection(ConnectionString);

        /// <summary>Test de connectivité léger (ouvre puis ferme).</summary>
        public static async Task<bool> EstJoignableAsync()
        {
            try { using var c = CreerConnexion(); await c.OpenAsync(); return true; }
            catch { return false; }
        }

        /// <summary>
        /// Diagnostic au démarrage : se connecte, journalise le serveur + le NOM RÉEL de la base
        /// par défaut, et vérifie la présence des tables Metrologo attendues. Best-effort (ne lève
        /// jamais). Le résultat va dans le Journal (catégorie Système) → on voit tout sans rien
        /// copier-coller.
        /// </summary>
        public static async Task DiagnostiquerAsync()
        {
            try
            {
                using var c = CreerConnexion();
                await c.OpenAsync();

                var row = await c.QueryFirstAsync("SELECT @@SERVERNAME AS Srv, DB_NAME() AS Db");
                string srv = (string)row.Srv;
                string db = (string)row.Db;

                var attendues = new[] { "T_METROLOGO_DATESRUBIS", "TJ_METROLOGO_SUIVIRUBI", "TR_METROLOGO_RUBIDIUMS" };
                List<string> presentes = (await c.QueryAsync<string>(
                    "SELECT name FROM sys.tables WHERE name IN @noms", new { noms = attendues })).ToList();
                var manquantes = attendues.Except(presentes, StringComparer.OrdinalIgnoreCase).ToList();

                JournalLog.Info(CategorieLog.Systeme, "DB_METROLOGO_OK",
                    $"Base Metrologo joignable — serveur « {srv} », base « {db} ». "
                  + $"Tables présentes : {(presentes.Count == 0 ? "aucune" : string.Join(", ", presentes))}"
                  + (manquantes.Count > 0 ? $" · MANQUANTES : {string.Join(", ", manquantes)}." : "."));
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Systeme, "DB_METROLOGO_KO",
                    $"Connexion base Metrologo (SVR-OR / MetrologoUser) échouée : {ex.Message}");
            }
        }
    }
}
