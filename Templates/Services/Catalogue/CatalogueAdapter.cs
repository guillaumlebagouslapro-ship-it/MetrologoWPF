using System.Globalization;
using System.Text.RegularExpressions;
using Metrologo.Models;

namespace Metrologo.Services.Catalogue
{
    /// <summary>
    /// Convertit un <see cref="ModeleAppareil"/> du catalogue en <see cref="AppareilIEEE"/> pilotable
    /// par <see cref="MesureOrchestrator"/> (n'importe quel appareil enregistré, par ex. un Agilent
    /// 53131A, même absent de l'ancien enum Stanford/Racal/EIP). L'adresse GPIB vient de la détection
    /// sur le bus, pas du catalogue, qui reste ainsi portable d'un banc à l'autre.
    /// </summary>
    public static class CatalogueAdapter
    {
        /// <summary>
        /// Construit l'<see cref="AppareilIEEE"/> à partir du modèle et de l'adresse trouvée sur le bus.
        /// Les gates (libellés "100 ms", "1 s"…) sont injectées dans le template <c>CommandeGate</c>
        /// du modèle (ex. <c>:FREQ:ARM:STOP:TIM {0}</c>).
        /// </summary>
        public static AppareilIEEE VersAppareilIEEE(ModeleAppareil modele, int adresse)
        {
            var p = modele.Parametres;

            var appareil = new AppareilIEEE
            {
                Nom = modele.Nom,
                Adresse = adresse,
                WriteTerm = p.TermWrite,
                ReadTerm = p.TermRead,
                TailleHeaderReponse = p.TailleHeader,
                ChaineInit = p.ChaineInit ?? string.Empty,
                ConfEntree = p.ConfEntree ?? string.Empty,
                ExeMesure = p.ExeMesure ?? string.Empty,
                Monocoup = string.Empty,
                CommandeMesureMultiple = p.CommandeMesureMultiple ?? string.Empty,
                CommandeFetchFresh = p.CommandeFetchFresh ?? string.Empty,
                GereSRQ = p.GereSrq,
                SRQOn = p.SrqOn ?? string.Empty,
                SRQOff = p.SrqOff ?? string.Empty,
                VerifArmingActive = p.VerifArmingActive,
                ModeRapideActif = p.ModeRapideActif,
                CommandeBulkInit = p.CommandeBulkInit ?? string.Empty,
                Gates = ConstruireGates(modele.Gates, p.CommandeGate, p.CommandesGateParSlot)
            };

            return appareil;
        }

        // Échelle de référence de l'UI (ConfigurationViewModel.GateTimes, SelectionGateViewModel.GateTimes)
        // et de ConfigAppareilsLoader._gates pour les appareils legacy. Les gates du catalogue doivent
        // se ranger dans ces mêmes slots, sinon Mesure.GateIndex ne retombe pas sur la bonne entrée.
        private static readonly double[] _secondesSlotsUi =
        {
            0.010, 0.020, 0.050, 0.100, 0.200, 0.500,
            1.0, 2.0, 5.0, 10.0, 20.0, 50.0,
            100.0, 200.0, 500.0, 1000.0 
        };

        /// <summary>
        /// Range des libellés (<c>["10 ms", "100 ms", "1 s"]</c>) dans un dictionnaire indexé par les
        /// slots de l'échelle UI ("1 s" → slot 6, "10 s" → slot 9…), comme les appareils legacy :
        /// sans cet alignement, <c>AppliquerGateAsync</c> ne retrouve pas la gate. Si
        /// <c>CommandeGate</c> est vide, la commande reste vide (l'orchestrator saute la
        /// programmation, cas du Racal en mode Interval).
        /// </summary>
        private static Dictionary<int, GateConfig> ConstruireGates(
            List<string> libellesGates, string templateCommande,
            Dictionary<int, string>? commandesParSlot = null)
        {
            var dict = new Dictionary<int, GateConfig>();
            foreach (var libelle in libellesGates)
            {
                double secondes = ParserGateEnSecondes(libelle);
                int slot = TrouverSlotUi(secondes);
                if (slot < 0) continue;  // Valeur hors standard : impossible à caser dans l'UI.

                // Appareils legacy : commande figée par slot (pas de template possible), prioritaire
                // sur le template {0} des appareils modernes.
                string commande;
                if (commandesParSlot != null && commandesParSlot.TryGetValue(slot, out var cmdLegacy)
                    && !string.IsNullOrWhiteSpace(cmdLegacy))
                    commande = cmdLegacy;
                else if (!string.IsNullOrWhiteSpace(templateCommande))
                    commande = string.Format(CultureInfo.InvariantCulture, templateCommande, secondes);
                else
                    commande = string.Empty;

                dict[slot] = new GateConfig
                {
                    Libelle = libelle,
                    Commande = commande,
                    ValeurSecondes = secondes
                };
            }
            return dict;
        }

        private static int TrouverSlotUi(double secondes)
        {
            for (int i = 0; i < _secondesSlotsUi.Length; i++)
            {
                if (Math.Abs(_secondesSlotsUi[i] - secondes) < 1e-6) return i;
            }
            return -1;
        }

        /// <summary>
        /// Convertit en secondes un libellé "10 ms", "1 s", "100 s" (espace facultatif). 1.0 par défaut
        /// si le format n'est pas reconnu.
        /// </summary>
        private static double ParserGateEnSecondes(string libelle)
        {
            if (string.IsNullOrWhiteSpace(libelle)) return 1.0;

            var match = Regex.Match(libelle.Trim(),
                @"^(?<n>[\d.,]+)\s*(?<u>ms|s)$",
                RegexOptions.IgnoreCase);
            if (!match.Success) return 1.0;

            string nbStr = match.Groups["n"].Value.Replace(',', '.');
            if (!double.TryParse(nbStr, NumberStyles.Float, CultureInfo.InvariantCulture, out var n))
                return 1.0;

            return match.Groups["u"].Value.ToLowerInvariant() == "ms" ? n / 1000.0 : n;
        }
    }
}
