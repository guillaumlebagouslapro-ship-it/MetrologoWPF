using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Ivi.Visa;
using Metrologo.Services.Journal;
using NationalInstruments.Visa;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Ieee
{
    /// <summary>
    /// Pilote IEEE réel basé sur NI-VISA via <c>NationalInstruments.Visa</c>.
    /// Convertit les primitives Delphi (EcritureIEEE, LectureIEEE, etc.) en sessions VISA.
    ///
    /// Les sessions GPIB sont mises en cache par adresse pour éviter de les rouvrir à chaque commande.
    /// Libérer via <see cref="Dispose"/> à la fermeture de l'application.
    /// </summary>
    public sealed class VisaIeeeDriver : IIeeeDriver, IDisposable
    {
        private readonly ResourceManager _rm;
        private readonly Dictionary<int, GpibSession> _sessions = new();
        private readonly int _gpibBoard;
        private readonly object _lock = new();
        private bool _disposed;

        /// <param name="gpibBoard">Index de la carte GPIB — 0 pour GPIB0 (cas général).</param>
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
            catch (Exception) { /* best-effort — certains backends refusent IFC en runtime */ }
            return Task.CompletedTask;
        }

        public Task EcrireAsync(int adresse, string commande, int writeTerm, CancellationToken ct = default)
        {
            EnsureNotDisposed();
            ct.ThrowIfCancellationRequested();
            var session = ObtenirSession(adresse);
            AppliquerTerminateurs(session, readTerm: null);

            // Respect de WriteTerm (cf. Metrologo.ini / catalogue) :
            //   0 = rien (pas de LF, pas d'EOI)
            //   1 = NL (LF) en fin de commande + EOI — convention Delphi la plus courante
            //   2 = EOI uniquement (pas de LF)
            session.SendEndEnabled = writeTerm != 0;
            string aEcrire = writeTerm == 1 ? commande + "\n" : commande;

            // Synchrone direct sur le thread d'orchestration (qui tourne déjà sur un thread pool
            // hors UI). Évite l'overhead Task.Run/await (~10-20 ms par opération) qui se voit
            // énormément quand on enchaîne 30 :FETCh? en mode rapide stabilité.
            try
            {
                session.RawIO.Write(aEcrire);
            }
            catch (IOTimeoutException)
            {
                // Timeout en écriture = l'appareil refuse la commande parce qu'il a typiquement
                // encore une réponse en attente dans son buffer de sortie. Device Clear pour
                // remettre l'appareil dans un état propre, puis on relaie l'exception.
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

            // Synchrone direct (cf. EcrireAsync). Le thread d'orchestration est dédié, on peut
            // le bloquer pendant l'IO GPIB sans impact UI.
            try
            {
                return Task.FromResult(session.RawIO.ReadString());
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
                session.Clear();  // SDC : Selected Device Clear (IEEE-488.2)
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
        /// Ferme les sessions GPIB en cache sans toucher au <c>ResourceManager</c>. Les prochaines
        /// opérations en rouvriront de fraîches — utile si un appareil a été éteint/rallumé ou
        /// que le bus est dans un état bancal après un timeout.
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
                        kv.Value.Clear();   // SDC : débloque le ReadString en cours côté instrument.
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

        // ---------------- Interne ----------------

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
        /// ReadTerm : 10 = LF, 13 = CR, 256 = STOPEnd (EOI uniquement).
        /// WriteTerm est géré par VISA via le mode EOI par défaut, donc on ne le configure pas activement.
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
