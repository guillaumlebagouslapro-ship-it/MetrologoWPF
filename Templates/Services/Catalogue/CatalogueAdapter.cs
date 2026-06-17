using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Metrologo.Models;

namespace Metrologo.Services.Catalogue
{
    /// <summary>
    /// Fait le pont entre un <see cref="ModeleAppareil"/> du catalogue local et un
    /// <see cref="AppareilIEEE"/> que <see cref="MesureOrchestrator"/> sait piloter. Grâce à ça,
    /// l'orchestrator peut piloter n'importe quel appareil que l'utilisateur a enregistré
    /// (un Agilent 53131A, par exemple) même s'il n'apparaît pas dans le vieil enum
    /// Stanford/Racal/EIP.
    ///
    /// L'adresse GPIB, elle, vient de la détection sur le bus et pas du catalogue : on veut
    /// que le catalogue reste portable d'un poste à l'autre, sachant qu'un même modèle peut
    /// très bien se retrouver à des adresses différentes selon le banc.
    /// </summary>
    public static class CatalogueAdapter
    {
        /// <summary>
        /// Construit l'<see cref="AppareilIEEE"/> à partir du modèle catalogue et de l'adresse
        /// trouvée sur le bus. Pour les gates, on lit les libellés en texte (par ex. "100 ms",
        /// "1 s") et on injecte la valeur dans le template <c>CommandeGate</c> du modèle
        /// (du genre <c>:FREQ:ARM:STOP:TIM {0}</c>).
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

        // L'échelle de référence 0..12, celle que l'UI utilise partout (ConfigurationViewModel.GateTimes,
        // SelectionGateViewModel.GateTimes) et que ConfigAppareilsLoader._gates utilise pour les
        // appareils legacy. Les gates du catalogue doivent absolument se ranger dans ces mêmes
        // slots : sinon le Mesure.GateIndex qui arrive de l'UI ne retombe pas sur la bonne entrée.
        private static readonly double[] _secondesSlotsUi =
        {
            0.010, 0.020, 0.050, 0.100, 0.200, 0.500,
            1.0, 2.0, 5.0, 10.0, 20.0, 50.0,
            100.0, 200.0, 500.0, 1000.0 
        };

        /// <summary>
        /// Prend la liste de libellés (par ex. <c>["10 ms", "100 ms", "1 s"]</c>) et la range dans un
        /// dictionnaire dont les clés sont les indices de l'échelle UI à 13 slots (0=10 ms … 12=100 s).
        /// Un "1 s" du catalogue atterrit dans le slot 6, "10 s" dans le slot 9, etc. — exactement
        /// comme les appareils legacy. Si on ne fait pas cet alignement, l'index que la UI envoie
        /// ne retrouve pas les clés du dict et <c>AppliquerGateAsync</c> passe à côté de la gate.
        /// Quand <c>CommandeGate</c> est vide, on laisse la commande vide elle aussi (l'orchestrator
        /// sautera la programmation — c'est le cas du Racal en mode Interval, par exemple).
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

                // Cas legacy : une commande de gate figée pour chaque slot (pas de template
                // possible), qui prend le pas sur le template {0}. À défaut, on retombe sur le
                // template (les appareils modernes).
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
        /// Convertit en secondes une chaîne du genre "10 ms", "1 s", "100 s". On accepte <c>ms</c>
        /// comme <c>s</c>, avec ou sans espace. Si on ne reconnaît rien, on renvoie 1.0 par défaut.
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
