using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Metrologo.Models;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Ieee
{
    /// <summary>
    /// Envoie à un appareil les commandes SCPI associées à un ensemble de réglages choisis
    /// dans la fenêtre Configuration. Toutes les commandes sont journalisées pour faciliter
    /// le diagnostic (on voit exactement ce qui part sur le bus).
    ///
    /// Utilise le driver VISA injecté. Si <c>commandes</c> est vide ou que toutes sont vides,
    /// ne fait rien.
    /// </summary>
    public static class AppareilScpiService
    {
        /// <summary>
        /// Envoie les commandes à l'adresse donnée, une par une, en respectant le terminateur
        /// de l'appareil (<see cref="ParametresIeee.TermWrite"/>).
        /// </summary>
        /// <param name="modele">Modèle source — utilisé pour récupérer <c>TermWrite</c> et journaliser.</param>
        /// <param name="adresse">Adresse primaire GPIB de l'appareil.</param>
        /// <param name="commandes">Commandes SCPI à envoyer dans l'ordre. Les chaînes vides sont ignorées.</param>
        /// <param name="driver">Driver IEEE (<see cref="VisaIeeeDriver"/> en prod, simulation en test).</param>
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
            }
        }
    }
}
