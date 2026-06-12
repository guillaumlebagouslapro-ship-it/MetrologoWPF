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
        /// <summary>
        /// Envoie les commandes une par une à l'adresse donnée, avec le TermWrite du modèle.
        /// Les chaînes vides sont ignorées.
        /// </summary>
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

                // le 53131A (et d'autres compteurs lents) a besoin de ~50 ms entre les commandes,
                // sinon certaines sont rejetées en silence et l'instrument peut rester dans un
                // état bancal qui bloque les sessions VISA suivantes (ex: le *RST de l'orchestrator
                // au début d'une mesure)
                await Task.Delay(50, ct);
            }

            // Délai final avant que la session VISA ne soit libérée (le caller utilise
            // souvent un `using var driver`). Évite que l'instrument soit encore en train
            // de traiter la dernière commande quand on rouvre une session sur la même
            // adresse pour la mesure (cause typique de timeout sur le premier :READ?).
            await Task.Delay(200, ct);
        }
    }
}
