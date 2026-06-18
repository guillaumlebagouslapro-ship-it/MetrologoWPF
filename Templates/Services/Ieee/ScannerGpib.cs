using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Ivi.Visa;
using Metrologo.Models;
using NationalInstruments.Visa;

namespace Metrologo.Services.Ieee
{
    /// <summary>Résultat du scan d'une adresse GPIB unique.</summary>
    public class ResultatScanGpib : INotifyPropertyChanged
    {
        public int Board { get; init; }
        public int Adresse { get; init; }
        public string Ressource { get; init; } = string.Empty;

        /// <summary>Interface par laquelle l'appareil a répondu (GPIB ou LAN).</summary>
        public TypeBus TypeBus { get; init; } = TypeBus.Gpib;

        /// <summary>Hôte/IP d'un appareil LAN (null en GPIB).</summary>
        public string? Hote { get; init; }

        public bool Repond { get; init; }
        public string? ReponseIdn { get; init; }
        public string? Erreur { get; init; }

        public string? Fabricant { get; init; }
        public string? Modele { get; init; }
        public string? NumeroSerie { get; init; }
        public string? Firmware { get; init; }

        public bool AErreur => !string.IsNullOrEmpty(Erreur);

        /// <summary>
        /// IDN incohérent = possible conflit d'adresse (réponses mélangées sur le bus).
        /// Best-effort : deux appareils identiques peuvent passer inaperçus.
        /// </summary>
        public bool ConflitAdressePossible => Repond && ScannerGpib.EstIdnSuspect(ReponseIdn);

        /// <summary>Étiquette courte : "GPIB0::15" ou "LAN 192.168.1.50".</summary>
        public string AdresseCourte => TypeBus == TypeBus.Lan
            ? $"LAN {Hote}"
            : $"GPIB{Board}::{Adresse}";

        public string Affichage => Repond
            ? $"{AdresseCourte} → {Fabricant} {Modele} (série {NumeroSerie}, fw {Firmware})"
            : $"{AdresseCourte} → pas de réponse";

        // ---- Reconnaissance catalogue (settable, mis à jour par le VM après chaque scan) ----

        private ModeleAppareil? _modeleCatalogue;
        public ModeleAppareil? ModeleCatalogue
        {
            get => _modeleCatalogue;
            set
            {
                if (ReferenceEquals(_modeleCatalogue, value)) return;
                _modeleCatalogue = value;
                Raise(nameof(ModeleCatalogue));
                Raise(nameof(EstEnregistre));
                Raise(nameof(PeutEtreEnregistre));
                Raise(nameof(NomAffiche));
            }
        }

        public bool EstEnregistre => ModeleCatalogue != null;
        public bool PeutEtreEnregistre => Repond && ModeleCatalogue == null;

        /// <summary>Nom à afficher : nom du catalogue si reconnu, sinon modèle IDN.</summary>
        public string NomAffiche => ModeleCatalogue?.Nom ?? Modele ?? "Inconnu";

        public event PropertyChangedEventHandler? PropertyChanged;
        private void Raise([CallerMemberName] string? n = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n ?? string.Empty));
    }

    /// <summary>
    /// Scanner du bus GPIB : balaie les adresses 1..30, envoie *IDN? et parse la réponse.
    /// Passe par NI-VISA (NationalInstruments.Visa).
    /// </summary>
    public static class ScannerGpib
    {
        // plage d'adresses scannées (0 = le contrôleur, on ne le scanne pas)
        public const int AdressePremiere = 1;
        public const int AdresseDerniere = 30;

        // extrait board + adresse primaire d'une chaîne VISA "GPIBx::y::INSTR"
        private static readonly Regex _regexGpib =
            new(@"^GPIB(\d+)::(\d+)(?:::\d+)?::INSTR$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // extrait l'hôte/IP d'une chaîne VISA TCPIP (ex "TCPIP0::192.168.1.50::inst0::INSTR")
        private static readonly Regex _regexTcpip =
            new(@"^TCPIP\d*::([^:]+)::", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Scanne le bus via ResourceManager.Find() (rapide). Si Find() ne donne rien, fallback
        /// en balayage séquentiel 1..30 sur le board indiqué (appareil pas encore vu par VISA).
        /// </summary>
        public static async Task<List<ResultatScanGpib>> ScannerAsync(
            int gpibBoard = 0,
            int timeoutMs = 2000,
            IProgress<ResultatScanGpib>? progress = null,
            CancellationToken ct = default)
        {
            var resultats = new List<ResultatScanGpib>();
            using var rm = new ResourceManager();

            // ---- Bus GPIB : Find() VISA, sinon fallback balayage séquentiel 1..30 ----
            List<string> ressourcesGpib;
            try { ressourcesGpib = new List<string>(rm.Find("GPIB?*INSTR")); }
            catch (NativeVisaException) { ressourcesGpib = new List<string>(); }

            if (ressourcesGpib.Count > 0)
            {
                foreach (var res in ressourcesGpib)
                {
                    ct.ThrowIfCancellationRequested();
                    var r = await Task.Run(() => InterrogerRessource(rm, res, timeoutMs), ct);
                    resultats.Add(r);
                    progress?.Report(r);
                }
            }
            else
            {
                for (int addr = AdressePremiere; addr <= AdresseDerniere; addr++)
                {
                    ct.ThrowIfCancellationRequested();
                    var r = await Task.Run(() => InterrogerAdresse(rm, gpibBoard, addr, timeoutMs), ct);
                    resultats.Add(r);
                    progress?.Report(r);
                }
            }

            // ---- Bus LAN (LXI) : on récupère les ressources TCPIP via Find("?*") (toutes
            //      interfaces) puis on filtre sur "TCPIP". Find("?*") est celui de « Lister les
            //      ressources VISA » : il remonte aussi bien le canal VXI-11 (inst0::INSTR) que le
            //      socket SCPI brut (port 5025, ::SOCKET), là où Find("TCPIP?*INSTR") seul oublie
            //      les sockets. C'est essentiel en liaison directe (link-local 169.254.x.x) : le
            //      VXI-11 y est souvent injoignable et le nom mDNS « xxx.local » ne se résout pas
            //      pour une socket brute — seule la ressource SOCKET (IP numérique) répond alors.
            //      Pas de balayage d'IP : VISA ne remonte que ce qu'il découvre ou ce qui est dans NI-MAX.
            List<string> toutesRessources;
            try { toutesRessources = new List<string>(rm.Find("?*")); }
            catch (NativeVisaException) { toutesRessources = new List<string>(); }

            var lan = new List<ResultatScanGpib>();
            foreach (var res in toutesRessources
                         .Where(r => r.StartsWith("TCPIP", StringComparison.OrdinalIgnoreCase))
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                ct.ThrowIfCancellationRequested();
                lan.Add(await Task.Run(() => InterrogerRessource(rm, res, timeoutMs), ct));
            }

            // Un même appareil expose souvent VXI-11 ET socket. On garde les canaux qui répondent,
            // dédupliqués par identité IDN (un appareil physique = une seule ligne ; VXI-11 préféré
            // en cas de double réponse). Si RIEN ne répond, on affiche quand même les ressources
            // découvertes (avec leur erreur) pour rester diagnosticable au lieu de masquer le LAN.
            var aGarder = lan.Where(x => x.Repond)
                         .GroupBy(x => $"{x.Fabricant}|{x.Modele}|{x.NumeroSerie}", StringComparer.OrdinalIgnoreCase)
                         .Select(g => g.OrderBy(x => EstSocket(x.Ressource) ? 1 : 0).First())
                         .ToList();
            if (aGarder.Count == 0)
                aGarder = lan;

            foreach (var r in aGarder)
            {
                resultats.Add(r);
                progress?.Report(r);
            }

            return resultats;
        }

        /// <summary>Scanne une seule adresse (test ciblé).</summary>
        public static Task<ResultatScanGpib> InterrogerAsync(
            int adresse, int gpibBoard = 0, int timeoutMs = 1000, CancellationToken ct = default)
        {
            return Task.Run(() =>
            {
                using var rm = new ResourceManager();
                return InterrogerAdresse(rm, gpibBoard, adresse, timeoutMs);
            }, ct);
        }

        // ---------------- Interne ----------------

        private static ResultatScanGpib InterrogerAdresse(
            ResourceManager rm, int gpibBoard, int addr, int timeoutMs)
        {
            string resource = $"GPIB{gpibBoard}::{addr}::INSTR";
            return InterrogerRessource(rm, resource, timeoutMs);
        }

        private static ResultatScanGpib InterrogerRessource(
            ResourceManager rm, string resource, int timeoutMs)
        {
            // extrait board + adresse de la chaîne VISA (ex "GPIB9::15::INSTR")
            int board = 0, addr = 0;
            var m = _regexGpib.Match(resource);
            if (m.Success)
            {
                int.TryParse(m.Groups[1].Value, out board);
                int.TryParse(m.Groups[2].Value, out addr);
            }

            bool estLan = resource.StartsWith("TCPIP", StringComparison.OrdinalIgnoreCase);
            TypeBus typeBus = estLan ? TypeBus.Lan : TypeBus.Gpib;

            // GPIB : interrogation directe de la ressource découverte (l'EOI gère la fin de ligne).
            if (!estLan)
                return TenterIdn(rm, resource, timeoutMs, terminateurLf: false, board, addr, typeBus, hote: null);

            // LAN : hôte/IP d'après la chaîne VISA découverte.
            string? hote = null;
            var ml = _regexTcpip.Match(resource);
            if (ml.Success) hote = ml.Groups[1].Value;

            // Ressource socket déjà découverte (port 5025) : interrogation directe en mode socket
            // (terminateur LF, pas d'EOI). Pas la peine de tenter le VXI-11, ce n'est pas ce canal.
            if (EstSocket(resource))
                return TenterIdn(rm, resource, timeoutMs, terminateurLf: true, board, addr, typeBus, hote);

            // On tente d'abord le canal découvert (VXI-11, inst0::INSTR), qui supporte tout
            // l'IEEE-488 (Device Clear, status byte). S'il ne répond pas — fréquent quand le
            // VXI-11 est bloqué par un pare-feu ou capricieux en link-local — on bascule sur le
            // socket SCPI brut (port standard 5025), une simple connexion TCP. On mémorise la
            // ressource qui a effectivement répondu : la mesure ouvrira le même canal. Ainsi un
            // appareil dont le VXI-11 marche garde le VXI-11, sans rien coder en dur par modèle.
            var viaVxi = TenterIdn(rm, resource, System.Math.Min(timeoutMs, 1500),
                terminateurLf: false, board, addr, typeBus, hote);
            if (viaVxi.Repond) return viaVxi;

            if (hote != null)
            {
                var viaSocket = TenterIdn(rm, $"TCPIP0::{hote}::{PortSocketScpi}::SOCKET",
                    timeoutMs, terminateurLf: true, board, addr, typeBus, hote);
                if (viaSocket.Repond) return viaSocket;
            }

            return viaVxi;   // échec des deux canaux : on renvoie le 1er résultat (avec son erreur)
        }

        // Port "SCPI raw socket" standard des instruments modernes (Keysight, R&S, Tektronix...).
        // Pourra devenir un champ du catalogue si un appareil utilise un autre port.
        private const int PortSocketScpi = 5025;

        // Une ressource VISA socket brut se termine par "::SOCKET" (ex "TCPIP0::169.254.1.2::5025::SOCKET").
        private static bool EstSocket(string resource)
            => resource.EndsWith("SOCKET", StringComparison.OrdinalIgnoreCase);

        /// <summary>Ouvre une ressource VISA, envoie *IDN? et construit le résultat. terminateurLf =
        /// vrai pour un socket (pas d'EOI : la fin de ligne se repère sur le LF).</summary>
        private static ResultatScanGpib TenterIdn(
            ResourceManager rm, string resource, int timeoutMs, bool terminateurLf,
            int board, int addr, TypeBus typeBus, string? hote)
        {
            try
            {
                using var session = (IMessageBasedSession)rm.Open(resource);
                session.TimeoutMilliseconds = timeoutMs;
                if (terminateurLf)
                {
                    session.TerminationCharacter = (byte)'\n';
                    session.TerminationCharacterEnabled = true;
                }

                session.FormattedIO.WriteLine("*IDN?");
                string reponse = session.FormattedIO.ReadLine()?.Trim() ?? string.Empty;

                var (fab, mod, ser, fw) = ParserIdn(reponse);
                return new ResultatScanGpib
                {
                    Board = board,
                    Adresse = addr,
                    Ressource = resource,
                    TypeBus = typeBus,
                    Hote = hote,
                    Repond = !string.IsNullOrWhiteSpace(reponse),
                    ReponseIdn = reponse,
                    Fabricant = fab,
                    Modele = mod,
                    NumeroSerie = ser,
                    Firmware = fw
                };
            }
            catch (IOTimeoutException)
            {
                return new ResultatScanGpib { Board = board, Adresse = addr, Ressource = resource, TypeBus = typeBus, Hote = hote, Repond = false };
            }
            catch (NativeVisaException ex)
            {
                return new ResultatScanGpib
                {
                    Board = board, Adresse = addr, Ressource = resource, TypeBus = typeBus, Hote = hote,
                    Repond = false,
                    Erreur = $"VISA {ex.ErrorCode}: {ex.Message.Split('\n')[0]}"
                };
            }
            catch (Exception ex)
            {
                return new ResultatScanGpib
                {
                    Board = board, Adresse = addr, Ressource = resource, TypeBus = typeBus, Hote = hote,
                    Repond = false,
                    Erreur = $"{ex.GetType().Name}: {ex.Message.Split('\n')[0]}"
                };
            }
        }

        /// <summary>Liste brute des ressources VISA (toutes interfaces). Utile pour vérifier
        /// que l'adaptateur voit les instruments, indépendamment du scan.</summary>
        public static Task<List<string>> ListerRessourcesAsync(CancellationToken ct = default)
        {
            return Task.Run(() =>
            {
                using var rm = new ResourceManager();
                try
                {
                    // "?*" = wildcard VISA pour "toutes les ressources"
                    return new List<string>(rm.Find("?*"));
                }
                catch (NativeVisaException ex) when (ex.ErrorCode == NativeErrorCode.ResourceNotFound)
                {
                    return new List<string>();
                }
            }, ct);
        }

        /// <summary>Parse une chaîne IDN IEEE-488.2 : Fabricant,Modèle,Série,Firmware.
        /// Ex : "HEWLETT-PACKARD,53131A,0,4613" ou "STANFORD,SR620,s/n 12345,ver 1.20".</summary>
        public static (string? fabricant, string? modele, string? serie, string? firmware) ParserIdn(string idn)
        {
            if (string.IsNullOrWhiteSpace(idn)) return (null, null, null, null);

            var parts = idn.Split(',');
            string? fab = parts.Length > 0 ? parts[0].Trim() : null;
            string? mod = parts.Length > 1 ? parts[1].Trim() : null;
            string? ser = parts.Length > 2 ? parts[2].Trim() : null;
            string? fw  = parts.Length > 3 ? parts[3].Trim() : null;

            return (fab, mod, ser, fw);
        }

        /// <summary>
        /// IDN suspect = plus de 4 champs ou caractères de contrôle (réponses superposées
        /// de deux appareils sur la même adresse). IDN normal : 4 champs, tout imprimable.
        /// </summary>
        public static bool EstIdnSuspect(string? idn)
        {
            if (string.IsNullOrWhiteSpace(idn)) return false;
            if (idn.Split(',').Length > 4) return true;
            foreach (char c in idn)
                if (char.IsControl(c) && c != '\t') return true;
            return false;
        }
    }
}
