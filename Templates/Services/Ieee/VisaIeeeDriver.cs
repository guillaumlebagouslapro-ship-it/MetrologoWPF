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
    /// Le vrai pilote IEEE, posé sur NI-VISA (NationalInstruments.Visa). On garde une session GPIB
    /// ouverte par adresse plutôt que d'en rouvrir une à chaque commande, et on fait le ménage
    /// (Dispose) quand l'appli se ferme.
    /// </summary>
    public sealed class VisaIeeeDriver : IIeeeDriver, IDisposable
    {
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
            catch (Exception) { /* on tente, sans plus : certains backends refusent l'IFC à chaud */ }
            return Task.CompletedTask;
        }

        public Task EcrireAsync(int adresse, string commande, int writeTerm, CancellationToken ct = default)
        {
            EnsureNotDisposed();
            ct.ThrowIfCancellationRequested();
            var session = ObtenirSession(adresse);
            AppliquerTerminateurs(session, readTerm: null);

            // Signification de WriteTerm (défini dans Metrologo.ini / le catalogue) :
            //   0 = rien (ni LF, ni EOI)
            //   1 = LF à la fin de la commande + EOI — c'est le cas le plus fréquent côté Delphi
            //   2 = EOI seul, sans LF
            session.SendEndEnabled = writeTerm != 0;
            string aEcrire = writeTerm == 1 ? commande + "\n" : commande;

            // Appel synchrone direct : on est déjà sur le thread d'orchestration (un thread du pool,
            // hors UI), donc inutile de payer le coût d'un Task.Run/await (~10-20 ms à chaque coup).
            // Ça se voit beaucoup quand on enchaîne 30 :FETCh? d'affilée en mode rapide stabilité.
            try
            {
                session.RawIO.Write(aEcrire);
            }
            catch (IOTimeoutException)
            {
                // Un timeout en écriture, ça veut généralement dire que l'appareil n'accepte pas la
                // commande parce qu'il a encore une réponse coincée dans son buffer de sortie. On lui
                // envoie un Device Clear pour le remettre d'aplomb, puis on relance l'exception.
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

            // Synchrone direct, même logique que dans EcrireAsync. Le thread d'orchestration nous
            // appartient, on peut donc le bloquer le temps de l'IO GPIB sans gêner l'UI.
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
        /// Referme les sessions en cache mais laisse le ResourceManager en place ; elles seront
        /// rouvertes au besoin. Pratique quand un appareil a été éteint puis rallumé, ou quand le
        /// bus est resté dans un état douteux après un timeout.
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
                        kv.Value.Clear();   // SDC : ça débloque le ReadString qui patiente côté instrument.
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
        /// ReadTerm : 10 = LF, 13 = CR, 256 = STOPEnd (c'est-à-dire EOI seul).
        /// Pour le WriteTerm on ne fait rien de spécial : VISA s'en occupe tout seul via son mode EOI
        /// par défaut.
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
