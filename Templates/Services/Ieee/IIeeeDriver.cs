using System.Threading;
using System.Threading.Tasks;

namespace Metrologo.Services.Ieee
{
    /// <summary>
    /// Pilote bas niveau du bus GPIB/IEEE-488 (une instance par poste, adaptateur USB-GPIB).
    /// Reprend les primitives Delphi (EcritureIEEE, LectureIEEE, ReadStatusByte...) sur GPIB0.
    /// Implémentations : SimulationIeeeDriver (sans matériel), VisaIeeeDriver (NI-VISA),
    /// Ni488Driver (ni4882.dll en direct).
    /// </summary>
    public interface IIeeeDriver
    {
        /// <summary>Envoie IFC (Interface Clear) sur le bus. À appeler avant une séquence.</summary>
        Task SendInterfaceClearAsync(CancellationToken ct = default);

        /// <summary>Écrit une commande sur l'appareil. writeTerm : 0=rien, 1=NL, 2=EOI (cf. Metrologo.ini).</summary>
        Task EcrireAsync(int adresse, string commande, int writeTerm, CancellationToken ct = default);

        /// <summary>
        /// Lit la réponse de l'appareil. readTerm : code ASCII du terminateur (10 = LF),
        /// ou 256 pour EOI. Retourne la réponse brute, chaîne vide sur timeout.
        /// </summary>
        Task<string> LireAsync(int adresse, int readTerm, CancellationToken ct = default);

        /// <summary>Serial poll : octet de statut IEEE-488.2 (bit 0x10 = MAV, bit 0x40 = RQS).</summary>
        Task<byte> LireStatusByteAsync(int adresse, CancellationToken ct = default);

        /// <summary>Positionne la ligne REN (Remote Enable) : true = remote, false = local.</summary>
        Task DefinirRemoteLocalAsync(int adresse, bool remote, CancellationToken ct = default);

        /// <summary>Ferme les sessions GPIB en cache (rouvertes à la demande). À appeler après un cycle
        /// alimentation ou si le bus est bloqué. Sans effet en simulation.</summary>
        void ReinitialiserSessions();

        /// <summary>
        /// SDC sur toutes les sessions ouvertes. Synchrone et thread-safe : appelé depuis
        /// "Arrêter la mesure" pendant qu'un autre thread est bloqué dans :FETCh?. Le SDC
        /// débloque l'instrument, la boucle de mesure voit le token annulé sans attendre la gate.
        /// </summary>
        void AborterToutesSessions();

        /// <summary>
        /// Ajuste le timeout de la session. À faire avant toute gate dépassant le timeout par
        /// défaut (gate 20 s → timeout 22 s mini) : sinon :READ? retourne vide et la réponse
        /// reste dans le buffer, bloquant le Write suivant.
        /// </summary>
        void DefinirTimeout(int adresse, int timeoutMs);
    }
}
