using System;
using System.Collections.Generic;
using System.ComponentModel;
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

            // ---- Bus LAN (LXI / VXI-11) : Find() des ressources TCPIP::INSTR. Pas de balayage
            //      séquentiel ici (on ne sonde pas une plage d'IP) : VISA ne remonte que les
            //      appareils qu'il découvre sur le sous-réseau ou configurés dans NI-MAX. ----
            List<string> ressourcesLan;
            try { ressourcesLan = new List<string>(rm.Find("TCPIP?*INSTR")); }
            catch (NativeVisaException) { ressourcesLan = new List<string>(); }

            foreach (var res in ressourcesLan)
            {
                ct.ThrowIfCancellationRequested();
                var r = await Task.Run(() => InterrogerRessource(rm, res, timeoutMs), ct);
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

            // type de bus + hôte LAN d'après la chaîne VISA
            bool estLan = resource.StartsWith("TCPIP", StringComparison.OrdinalIgnoreCase);
            TypeBus typeBus = estLan ? TypeBus.Lan : TypeBus.Gpib;
            string? hote = null;
            if (estLan)
            {
                var ml = _regexTcpip.Match(resource);
                if (ml.Success) hote = ml.Groups[1].Value;
            }

            try
            {
                using var session = (IMessageBasedSession)rm.Open(resource);
                session.TimeoutMilliseconds = timeoutMs;

                // FormattedIO gère les terminateurs
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
