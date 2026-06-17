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
    /// Pilote IEEE qui attaque directement le NI-488.2 natif (P/Invoke sur ni4882.dll). C'est
    /// l'alternative à VisaIeeeDriver : nettement plus rapide sur les cycles write+read courts
    /// (~30-80 ms au lieu de ~190 ms en VISA), en échange d'un peu de tuyauterie à la main pour
    /// gérer le statut et les erreurs. On retrouve l'approche du Delphi d'origine (dpib32).
    /// </summary>
    public sealed class Ni488Driver : IIeeeDriver, IDisposable
    {
        private readonly int _boardId;
        private readonly Dictionary<int, int> _handles = new();   // adresse → ud (handle NI)
        private readonly object _lock = new();
        private bool _disposed;

        // On réutilise le même tampon de lecture d'un cycle à l'autre, histoire de ne pas allouer à chaque fetch.
        private readonly byte[] _bufLecture = new byte[4096];

        public Ni488Driver(int boardId = 0) { _boardId = boardId; }

        // ---------------- Implémentation de IIeeeDriver ----------------

        public Task SendInterfaceClearAsync(CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            EnsureNotDisposed();
            try { Ni488Native.SendIFC(_boardId); }
            catch { /* on tente, tant pis si ça échoue */ }
            return Task.CompletedTask;
        }

        public Task EcrireAsync(int adresse, string commande, int writeTerm, CancellationToken ct = default)
        {
            ct.ThrowIfCancellationRequested();
            EnsureNotDisposed();
            int ud = ObtenirHandle(adresse);

            // writeTerm 1 = LF + EOI (la convention historique)
            // writeTerm 2 = EOI tout seul
            // writeTerm 0 = rien du tout (ni LF, ni EOI)
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

            // Règle le caractère EOS (le terminateur) attendu côté instrument
            //   readTerm 10 = LF, 13 = CR, 256 = EOI seul
            if (readTerm > 0 && readTerm < 256)
            {
                Ni488Native.ibeos(ud, Ni488Native.REOS | readTerm);
            }
            else
            {
                Ni488Native.ibeos(ud, 0); // aucun char EOS : on lit jusqu'à l'EOI
            }

            int sta = Ni488Native.ibrd(ud, _bufLecture, _bufLecture.Length);
            if ((sta & Ni488Native.ERR_BIT) != 0)
            {
                if ((sta & Ni488Native.TIMO_BIT) != 0)
                {
                    // Timeout en lecture : on envoie un Device Clear pour remettre la session d'aplomb.
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
            // Pas encore implémenté : la mesure actuelle n'en a pas besoin.
            // Si un jour il le faut : ibloc pour repasser en local, ou LLO en mode commande.
            return Task.CompletedTask;
        }

        public void ReinitialiserSessions()
        {
            if (_disposed) return;
            lock (_lock)
            {
                foreach (var ud in _handles.Values)
                {
                    try { Ni488Native.ibonl(ud, 0); } catch { /* on ferme au mieux, sans s'inquiéter d'un échec */ }
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

        // ---------------- Détails internes ----------------

        private int ObtenirHandle(int adresse)
        {
            lock (_lock)
            {
                if (_handles.TryGetValue(adresse, out var existant)) return existant;

                // ibdev (board, pad, sad=0, tmo=T10s, eot=1, eos=0)
                int ud = Ni488Native.ibdev(_boardId, adresse, 0,
                    tmo: 13,        // T10s pour commencer
                    eot: 1,         // EOI affirmé sur le dernier octet
                    eos: 0);        // aucun char EOS au départ

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
