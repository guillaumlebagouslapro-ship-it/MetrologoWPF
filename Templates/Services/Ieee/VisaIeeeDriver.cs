using Ivi.Visa;
using Metrologo.Services.Journal;
using NationalInstruments.Visa;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Ieee
{
    /// <summary>
    /// Pilote IEEE sur NI-VISA (NationalInstruments.Visa). Une session GPIB reste ouverte par
    /// adresse ; elles sont fermées au Dispose.
    /// </summary>
    public sealed class VisaIeeeDriver : IIeeeDriver, IDisposable
    {
        /// <summary>Taille d'un bloc de lecture VISA (une réponse plus longue est relue par blocs).</summary>
        private const long TAILLE_BLOC_LECTURE = 4096;
        /// <summary>Garde-fou contre une lecture sans fin (≈ 40 000 valeurs).</summary>
        private const int TAILLE_MAX_REPONSE = 1_000_000;

        private readonly ResourceManager _rm;
        private readonly Dictionary<int, GpibSession> _sessions = new();
        private readonly int _gpibBoard;
        private readonly object _lock = new();
        private bool _disposed;

        /// <summary>gpibBoard : numéro de la carte ; en pratique c'est presque toujours 0 (GPIB0).</summary>
        public VisaIeeeDriver(int gpibBoard = 0)
        {
            _rm = new ResourceManager();
            _gpibBoard = gpibBoard;
        }

        public Task SendInterfaceClearAsync(CancellationToken ct = default)
        {
            EnsureNotDisposed();
            string resource = $"GPIB{_gpibBoard}::INTFC";
            try
            {
                using var intf = (GpibInterface)_rm.Open(resource);
                intf.SendInterfaceClear();
            }
            catch (Exception) { /* au mieux : certains backends refusent l'IFC à chaud */ }
            return Task.CompletedTask;
        }

        public Task EcrireAsync(int adresse, string commande, int writeTerm, CancellationToken ct = default)
        {
            EnsureNotDisposed();
            ct.ThrowIfCancellationRequested();
            var session = ObtenirSession(adresse);
            AppliquerTerminateurs(session, readTerm: null);

            // WriteTerm (Metrologo.ini / catalogue) : 0 = rien (ni LF ni EOI), 1 = LF + EOI (cas le
            // plus fréquent côté Delphi), 2 = EOI seul.
            session.SendEndEnabled = writeTerm != 0;
            string aEcrire = writeTerm == 1 ? commande + "\n" : commande;

            // Appel synchrone : on est déjà sur le thread d'orchestration (hors UI), un Task.Run
            // coûterait ~10-20 ms par appel, sensible sur 30 :FETCh? enchaînés en mode rapide.
            try
            {
                session.RawIO.Write(aEcrire);
            }
            catch (IOTimeoutException)
            {
                // Timeout en écriture : en général une réponse coincée dans le buffer de sortie de
                // l'appareil. Device Clear pour le remettre d'aplomb, puis on relance.
                TenterDeviceClear(session, adresse, "EcrireAsync");
                throw;
            }
            return Task.CompletedTask;
        }

        public Task<string> LireAsync(int adresse, int readTerm, CancellationToken ct = default)
        {
            EnsureNotDisposed();
            ct.ThrowIfCancellationRequested();
            var session = ObtenirSession(adresse);
            AppliquerTerminateurs(session, readTerm: readTerm);

            // Synchrone, comme EcrireAsync (thread d'orchestration, hors UI).
            // PIÈGE : RawIO.ReadString() sans argument lit au plus 1024 octets. Une réponse plus
            // longue (lot de 100 valeurs, ~2 300 car.) resterait en partie dans le buffer du compteur
            // et la commande suivante donnerait -410 Query INTERRUPTED. On relit donc par blocs tant
            // que VISA signale « nombre max atteint » (ni terminateur ni EOI).
            try
            {
                var reponse = new System.Text.StringBuilder();
                ReadStatus statut;
                do
                {
                    reponse.Append(session.RawIO.ReadString(TAILLE_BLOC_LECTURE, out statut));
                }
                while (statut == ReadStatus.MaximumCountReached && reponse.Length < TAILLE_MAX_REPONSE);
                return Task.FromResult(reponse.ToString());
            }
            catch (IOTimeoutException)
            {
                TenterDeviceClear(session, adresse, "LireAsync");
                return Task.FromResult(string.Empty);
            }
        }

        private static void TenterDeviceClear(GpibSession session, int adresse, string origine)
        {
            try
            {
                session.Clear();  // SDC, le Selected Device Clear de l'IEEE-488.2
                JournalLog.Warn(CategorieLog.Mesure, "GPIB_TIMEOUT_SDC",
                    $"Timeout sur {origine} GPIB0::{adresse} — Device Clear envoyé pour ré-aligner la session.",
                    new { Adresse = adresse, Origine = origine });
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Mesure, "GPIB_TIMEOUT_SDC_ECHEC",
                    $"Timeout sur {origine} GPIB0::{adresse} + échec Device Clear : {ex.GetType().Name}.",
                    new { Adresse = adresse, Origine = origine, Erreur = ex.Message });
            }
        }

        public Task<byte> LireStatusByteAsync(int adresse, CancellationToken ct = default)
        {
            EnsureNotDisposed();
            var session = ObtenirSession(adresse);
            return Task.Run(() => (byte)session.ReadStatusByte(), ct);
        }

        public Task DefinirRemoteLocalAsync(int adresse, bool remote, CancellationToken ct = default)
        {
            EnsureNotDisposed();
            var session = ObtenirSession(adresse);
            return Task.Run(() =>
            {
                var mode = remote ? RemoteLocalMode.Remote : RemoteLocalMode.Local;
                session.SendRemoteLocalCommand(mode);
            }, ct);
        }

        public void Dispose()
        {
            if (_disposed) return;
            lock (_lock)
            {
                foreach (var s in _sessions.Values)
                {
                    try { s.Dispose(); } catch { /* ignore */ }
                }
                _sessions.Clear();
                try { _rm.Dispose(); } catch { /* ignore */ }
            }
            _disposed = true;
        }

        public void DefinirTimeout(int adresse, int timeoutMs)
        {
            EnsureNotDisposed();
            var session = ObtenirSession(adresse);
            lock (_lock)
            {
                session.TimeoutMilliseconds = timeoutMs;
            }
        }

        /// <summary>
        /// Ferme les sessions en cache (rouvertes au besoin), le ResourceManager reste en place.
        /// Utile après extinction/rallumage d'un appareil ou un bus douteux après timeout.
        /// </summary>
        public void ReinitialiserSessions()
        {
            if (_disposed) return;
            lock (_lock)
            {
                foreach (var s in _sessions.Values)
                {
                    try { s.Dispose(); } catch { /* best-effort */ }
                }
                _sessions.Clear();
            }
        }

        public void AborterToutesSessions()
        {
            if (_disposed) return;
            lock (_lock)
            {
                foreach (var kv in _sessions)
                {
                    try
                    {
                        kv.Value.Clear();   // SDC : débloque le ReadString en attente
                        JournalLog.Warn(CategorieLog.Mesure, "GPIB_ABORT_SDC",
                            $"Device Clear envoyé à GPIB0::{kv.Key} (arrêt utilisateur).");
                    }
                    catch (Exception ex)
                    {
                        JournalLog.Warn(CategorieLog.Mesure, "GPIB_ABORT_SDC_ECHEC",
                            $"Device Clear sur GPIB0::{kv.Key} échoué : {ex.GetType().Name} — {ex.Message}.");
                    }
                }
            }
        }

        // ---------------- Détails internes ----------------

        private GpibSession ObtenirSession(int adresse)
        {
            lock (_lock)
            {
                if (_sessions.TryGetValue(adresse, out var existante))
                    return existante;

                string resource = $"GPIB{_gpibBoard}::{adresse}::INSTR";
                var session = (GpibSession)_rm.Open(resource);
                session.TimeoutMilliseconds = 5000;
                _sessions[adresse] = session;
                return session;
            }
        }

        /// <summary>
        /// ReadTerm : 10 = LF, 13 = CR, 256 = STOPEnd (EOI seul). Le WriteTerm est géré dans EcrireAsync.
        /// </summary>
        private static void AppliquerTerminateurs(GpibSession session, int? readTerm)
        {
            if (!readTerm.HasValue) return;
            int rt = readTerm.Value;

            if (rt == 256)
            {
                session.TerminationCharacterEnabled = false;
            }
            else if (rt > 0 && rt < 256)
            {
                session.TerminationCharacter = (byte)rt;
                session.TerminationCharacterEnabled = true;
            }
        }

        private void EnsureNotDisposed()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(VisaIeeeDriver));
        }
    }
}
