using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Metrologo.Models;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Ieee
{
    /// <summary>
    /// Envoie à un appareil les commandes SCPI des réglages choisis dans la fenêtre Configuration.
    /// Tout est journalisé : on voit exactement ce qui part sur le bus.
    /// </summary>
    public static class AppareilScpiService
    {
        /// <summary>Envoie les commandes une par une avec le TermWrite du modèle. Chaînes vides ignorées.</summary>
        public static async Task EnvoyerAsync(
            ModeleAppareil modele,
            int adresse,
            IEnumerable<string> commandes,
            IIeeeDriver driver,
            CancellationToken ct = default)
        {
            int termWrite = modele.Parametres.TermWrite;

            foreach (var cmd in commandes)
            {
                if (string.IsNullOrWhiteSpace(cmd)) continue;

                JournalLog.Info(CategorieLog.Mesure, "SCPI_ENVOI",
                    $"GPIB0::{adresse} ← {cmd}",
                    new { modele.Nom, Adresse = adresse, Commande = cmd });

                await driver.EcrireAsync(adresse, cmd, termWrite, ct);

                // 53131A (et autres compteurs lents) : rejette silencieusement les commandes trop
                // rapprochées, ce qui laisse l'instrument dans un état bancal bloquant les sessions
                // VISA suivantes (ex: *RST en début de mesure).
                await Task.Delay(50, ct);
            }

            // Laisse l'instrument terminer la dernière commande avant que la session VISA soit
            // libérée (caller en `using`). Évite le timeout sur le premier :READ? de la mesure.
            await Task.Delay(200, ct);
        }
    }
}
