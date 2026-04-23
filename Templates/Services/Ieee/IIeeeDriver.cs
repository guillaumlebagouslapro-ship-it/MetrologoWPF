using System.Threading;
using System.Threading.Tasks;

namespace Metrologo.Services.Ieee
{
    /// <summary>
    /// Pilote bas niveau du bus GPIB/IEEE-488. Une seule instance par poste (adaptateur USB-GPIB).
    /// Portage direct des primitives Delphi (EcritureIEEE, LectureIEEE, ReadStatusByte, etc.)
    /// définies historiquement sur <c>GPIB0</c>.
    ///
    /// Implémentations :
    ///   - <see cref="SimulationIeeeDriver"/> : pas de matériel, pour dev/tests.
    ///   - VisaIeeeDriver (à venir) : NI-VISA via NationalInstruments.Visa.
    /// </summary>
    public interface IIeeeDriver
    {
        /// <summary>Envoie <c>IFC</c> (Interface Clear) sur le bus. À appeler avant une séquence.</summary>
        Task SendInterfaceClearAsync(CancellationToken ct = default);

        /// <summary>Écrit une commande sur l'appareil à l'adresse donnée.</summary>
        /// <param name="writeTerm">Terminateur en écriture (0=none, 1=NL, 2=EOI — cf. Metrologo.ini).</param>
        Task EcrireAsync(int adresse, string commande, int writeTerm, CancellationToken ct = default);

        /// <summary>Lit la réponse de l'appareil à l'adresse donnée.</summary>
        /// <param name="readTerm">Code ASCII du terminateur de lecture (ex: 10 = LF), ou 256 pour EOI.</param>
        /// <returns>Réponse brute, ou chaîne vide en cas de timeout.</returns>
        Task<string> LireAsync(int adresse, int readTerm, CancellationToken ct = default);

        /// <summary>Effectue un serial poll et retourne l'octet de statut IEEE-488.2.</summary>
        /// <remarks>Bit 0x10 (16) = MAV (Message Available). Bit 0x40 (64) = RQS (service request).</remarks>
        Task<byte> LireStatusByteAsync(int adresse, CancellationToken ct = default);

        /// <summary>Positionne la ligne REN (Remote Enable) — true = remote, false = local.</summary>
        Task DefinirRemoteLocalAsync(int adresse, bool remote, CancellationToken ct = default);

        /// <summary>
        /// Ferme toutes les sessions GPIB en cache — les prochaines écritures/lectures en rouvriront
        /// de fraîches. À appeler quand un appareil a été éteint/rallumé ou qu'on soupçonne un
        /// état bus GPIB bloqué. N'a pas d'effet pour un driver sans cache (simulation).
        /// </summary>
        void ReinitialiserSessions();
    }
}
