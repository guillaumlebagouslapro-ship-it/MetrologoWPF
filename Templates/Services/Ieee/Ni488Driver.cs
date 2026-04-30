using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Ieee
{
    /// <summary>
    /// Pilote IEEE basé sur NI-488.2 natif (P/Invoke direct sur ni4882.dll).
    /// Alternative à <see cref="VisaIeeeDriver"/> qui passe par la couche NI-VISA managée.
    /// Beaucoup plus rapide pour les cycles write+read courts (~30-80 ms vs ~190 ms en VISA),
    /// au prix d'un peu plus de boilerplate côté status/erreur. Approche identique au Delphi
    /// historique (qui utilisait dpib32, wrapper Pascal de la même DLL).
    /// </summary>
    public sealed class Ni488Driver : IIeeeDriver, IDisposable
    {
        private readonly int _boardId;
        private readonly Dictionary<int, int> _handles = new();   // adresse → ud
        private readonly object _lock = new();
        private bool _disposed;

        // Tampon de lecture réutilisé entre cycles (évite l'allocation par fetch).
        private readonly byte[] _bufLecture = new byte[4096];

        public Ni488Driver(int boardId = 0) { _boardId = boardId; }

        // ---------------- IIeeeDriver ----------------

        public Task SendInterfaceClearAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            EnsureNotDisposed();
            try { Ni488Native.SendIFC(_boardId); }
            catch { /* best-effort */ }
            return Task.CompletedTask;
        }

        public Task EcrireAsync(int adresse, string commande, int writeTerm, CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            EnsureNotDisposed();
            int ud = ObtenirHandle(adresse);

            // writeTerm 1 = LF + EOI (convention historique)
            // writeTerm 2 = EOI uniquement
            // writeTerm 0 = rien (pas de LF, pas d'EOI)
            string aEnvoyer = writeTerm == 1 ? commande + "\n" : commande;
            byte[] buf = Encoding.ASCII.GetBytes(aEnvoyer);

            int sta = Ni488Native.ibwrt(ud, buf, buf.Length);
            if ((sta & Ni488Native.ERR_BIT) != 0)
            {
                int err = Ni488Native.ThreadIberr();
                if ((sta & Ni488Native.TIMO_BIT) != 0)
                {
                    TenterDeviceClear(ud, adresse, "EcrireAsync");
                    throw new IOException(
                        $"GPIB write timeout sur adresse {adresse} (sta=0x{sta:X4}, err={err})");
                }
                throw new IOException(
                    $"GPIB write erreur sur adresse {adresse} (sta=0x{sta:X4}, err={err})");
            }
            return Task.CompletedTask;
        }

        public Task<string> LireAsync(int adresse, int readTerm, CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            EnsureNotDisposed();
            int ud = ObtenirHandle(adresse);

            // Configure le caractère EOS (terminateur) côté instrument
            //   readTerm 10 = LF, 13 = CR, 256 = EOI uniquement
            if (readTerm > 0 && readTerm < 256)
            {
                Ni488Native.ibeos(ud, Ni488Native.REOS | readTerm);
            }
            else
            {
                Ni488Native.ibeos(ud, 0); // pas de char EOS, lecture jusqu'à EOI
            }

            int sta = Ni488Native.ibrd(ud, _bufLecture, _bufLecture.Length);
            if ((sta & Ni488Native.ERR_BIT) != 0)
            {
                if ((sta & Ni488Native.TIMO_BIT) != 0)
                {
                    // Timeout en lecture : Device Clear pour ré-aligner la session.
                    TenterDeviceClear(ud, adresse, "LireAsync");
                    return Task.FromResult(string.Empty);
                }
                int err = Ni488Native.ThreadIberr();
                throw new IOException(
                    $"GPIB read erreur sur adresse {adresse} (sta=0x{sta:X4}, err={err})");
            }

            int count = Ni488Native.ThreadIbcntl();
            if (count <= 0) return Task.FromResult(string.Empty);

            string s = Encoding.ASCII.GetString(_bufLecture, 0, count);
            return Task.FromResult(s.TrimEnd('\r', '\n', '\0', ' '));
        }

        public Task<byte> LireStatusByteAsync(int adresse, CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            EnsureNotDisposed();
            int ud = ObtenirHandle(adresse);
            Ni488Native.ibrsp(ud, out byte spr);
            return Task.FromResult(spr);
        }

        public Task DefinirRemoteLocalAsync(int adresse, bool remote, CancellationToken ct = default)
        {
            // Non implémenté pour l'instant — pas utilisé par la mesure courante.
            // Possible via ibloc (local) ou en envoyant LLO en command mode.
            return Task.CompletedTask;
        }

        public void ReinitialiserSessions()
        {
            if (_disposed) return;
            lock (_lock)
            {
                foreach (var ud in _handles.Values)
                {
                    try { Ni488Native.ibonl(ud, 0); } catch { /* best-effort */ }
                }
                _handles.Clear();
            }
        }

        public void AborterToutesSessions()
        {
            if (_disposed) return;
            lock (_lock)
            {
                foreach (var kv in _handles)
                {
                    try
                    {
                        Ni488Native.ibclr(kv.Value);
                        JournalLog.Warn(CategorieLog.Mesure, "GPIB_ABORT_SDC",
                            $"Device Clear envoyé à GPIB0::{kv.Key} (arrêt utilisateur).");
                    }
                    catch (Exception ex)
                    {
                        JournalLog.Warn(CategorieLog.Mesure, "GPIB_ABORT_SDC_ECHEC",
                            $"Device Clear sur GPIB0::{kv.Key} échoué : {ex.Message}.");
                    }
                }
            }
        }

        public void DefinirTimeout(int adresse, int timeoutMs)
        {
            EnsureNotDisposed();
            int ud = ObtenirHandle(adresse);
            int code = Ni488Native.MapTimeoutCode(timeoutMs);
            Ni488Native.ibtmo(ud, code);
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            ReinitialiserSessions();
        }

        // ---------------- Interne ----------------

        private int ObtenirHandle(int adresse)
        {
            lock (_lock)
            {
                if (_handles.TryGetValue(adresse, out var existant)) return existant;

                // ibdev (board, pad, sad=0, tmo=T10s, eot=1, eos=0)
                int ud = Ni488Native.ibdev(_boardId, adresse, 0,
                    tmo: 13,        // T10s par défaut
                    eot: 1,         // assert EOI sur dernier octet
                    eos: 0);        // pas de char EOS au démarrage

                if (ud < 0)
                {
                    int sta = Ni488Native.ThreadIbsta();
                    int err = Ni488Native.ThreadIberr();
                    throw new IOException(
                        $"ibdev a échoué pour GPIB{_boardId}::{adresse} (sta=0x{sta:X4}, err={err})");
                }
                _handles[adresse] = ud;
                return ud;
            }
        }

        private void TenterDeviceClear(int ud, int adresse, string origine)
        {
            try
            {
                Ni488Native.ibclr(ud);
                JournalLog.Warn(CategorieLog.Mesure, "GPIB_TIMEOUT_SDC",
                    $"Timeout sur {origine} GPIB0::{adresse} — Device Clear envoyé pour ré-aligner la session.",
                    new { Adresse = adresse, Origine = origine });
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Mesure, "GPIB_TIMEOUT_SDC_ECHEC",
                    $"Timeout sur {origine} GPIB0::{adresse} + échec Device Clear : {ex.Message}.");
            }
        }

        private void EnsureNotDisposed()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(Ni488Driver));
        }
    }
}
