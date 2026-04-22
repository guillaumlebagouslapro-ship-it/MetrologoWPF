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
        public bool Repond { get; init; }
        public string? ReponseIdn { get; init; }
        public string? Erreur { get; init; }

        public string? Fabricant { get; init; }
        public string? Modele { get; init; }
        public string? NumeroSerie { get; init; }
        public string? Firmware { get; init; }

        public bool AErreur => !string.IsNullOrEmpty(Erreur);

        /// <summary>Étiquette courte, ex: <c>GPIB0::15</c>.</summary>
        public string AdresseCourte => $"GPIB{Board}::{Adresse}";

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
    /// Scanner du bus GPIB : balaie les adresses primaires 1..30, envoie <c>*IDN?</c>
    /// à chaque appareil qui répond, et parse la chaîne d'identification retournée.
    ///
    /// Utilise NI-VISA via <c>NationalInstruments.Visa</c>.
    /// </summary>
    public static class ScannerGpib
    {
        /// <summary>
        /// Plage d'adresses GPIB utilisables. L'adresse 0 est habituellement celle du contrôleur.
        /// </summary>
        public const int AdressePremiere = 1;
        public const int AdresseDerniere = 30;

        // Regex pour extraire board + adresse primaire d'une chaîne VISA "GPIB<b>::<a>::INSTR".
        private static readonly Regex _regexGpib =
            new(@"^GPIB(\d+)::(\d+)(?:::\d+)?::INSTR$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>
        /// Scanne le bus GPIB en s'appuyant sur <c>ResourceManager.Find()</c>. C'est beaucoup plus rapide
        /// que d'itérer toutes les adresses : NI-VISA sait déjà qui répond sur le bus.
        ///
        /// <para>Si <c>Find()</c> ne retourne rien, fallback sur un balayage séquentiel (1..30)
        /// sur le board indiqué — utile si un appareil est branché mais n'a pas encore été "annoncé" à VISA.</para>
        /// </summary>
        public static async Task<List<ResultatScanGpib>> ScannerAsync(
            int gpibBoard = 0,
            int timeoutMs = 2000,
            IProgress<ResultatScanGpib>? progress = null,
            CancellationToken ct = default)
        {
            var resultats = new List<ResultatScanGpib>();
            using var rm = new ResourceManager();

            // ---- Passe 1 : utilisation de Find() (méthode VISA native) ----
            List<string> ressources;
            try { ressources = new List<string>(rm.Find("GPIB?*INSTR")); }
            catch (NativeVisaException) { ressources = new List<string>(); }

            if (ressources.Count > 0)
            {
                foreach (var res in ressources)
                {
                    ct.ThrowIfCancellationRequested();
                    var r = await Task.Run(() => InterrogerRessource(rm, res, timeoutMs), ct);
                    resultats.Add(r);
                    progress?.Report(r);
                }
                return resultats;
            }

            // ---- Passe 2 (fallback) : balayage séquentiel sur le board indiqué ----
            for (int addr = AdressePremiere; addr <= AdresseDerniere; addr++)
            {
                ct.ThrowIfCancellationRequested();
                var r = await Task.Run(() => InterrogerAdresse(rm, gpibBoard, addr, timeoutMs), ct);
                resultats.Add(r);
                progress?.Report(r);
            }

            return resultats;
        }

        /// <summary>Scanne une seule adresse (utilise en test ciblé).</summary>
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
            // Extraction du board + adresse depuis la chaîne VISA (ex: "GPIB9::15::INSTR").
            int board = 0, addr = 0;
            var m = _regexGpib.Match(resource);
            if (m.Success)
            {
                int.TryParse(m.Groups[1].Value, out board);
                int.TryParse(m.Groups[2].Value, out addr);
            }

            try
            {
                using var session = (IMessageBasedSession)rm.Open(resource);
                session.TimeoutMilliseconds = timeoutMs;

                // FormattedIO pour gérer les terminateurs proprement.
                session.FormattedIO.WriteLine("*IDN?");
                string reponse = session.FormattedIO.ReadLine()?.Trim() ?? string.Empty;

                var (fab, mod, ser, fw) = ParserIdn(reponse);
                return new ResultatScanGpib
                {
                    Board = board,
                    Adresse = addr,
                    Ressource = resource,
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
                return new ResultatScanGpib { Board = board, Adresse = addr, Ressource = resource, Repond = false };
            }
            catch (NativeVisaException ex)
            {
                return new ResultatScanGpib
                {
                    Board = board, Adresse = addr, Ressource = resource,
                    Repond = false,
                    Erreur = $"VISA {ex.ErrorCode}: {ex.Message.Split('\n')[0]}"
                };
            }
            catch (Exception ex)
            {
                return new ResultatScanGpib
                {
                    Board = board, Adresse = addr, Ressource = resource,
                    Repond = false,
                    Erreur = $"{ex.GetType().Name}: {ex.Message.Split('\n')[0]}"
                };
            }
        }

        /// <summary>
        /// Retourne la liste brute des ressources VISA que NI connaît (toutes interfaces confondues).
        /// Utile pour vérifier que l'adaptateur voit bien les instruments, indépendamment de notre scan séquentiel.
        /// </summary>
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

        /// <summary>
        /// Parse une chaîne IDN standard IEEE-488.2 : <c>Fabricant,Modèle,Série,Firmware</c>.
        /// Exemples : <c>HEWLETT-PACKARD,53131A,0,4613</c> · <c>STANFORD,SR620,s/n 12345,ver 1.20</c>.
        /// </summary>
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
    }
}
