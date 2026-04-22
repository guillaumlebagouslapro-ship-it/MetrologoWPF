using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Ivi.Visa;
using NationalInstruments.Visa;

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
            var session = ObtenirSession(adresse);
            AppliquerTerminateurs(session, readTerm: null);

            return Task.Run(() => session.RawIO.Write(commande), ct);
        }

        public Task<string> LireAsync(int adresse, int readTerm, CancellationToken ct = default)
        {
            EnsureNotDisposed();
            var session = ObtenirSession(adresse);
            AppliquerTerminateurs(session, readTerm: readTerm);

            return Task.Run(() =>
            {
                try { return session.RawIO.ReadString(); }
                catch (IOTimeoutException) { return string.Empty; }
            }, ct);
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
