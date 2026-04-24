using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Metrologo.Models;

namespace Metrologo.Services.Catalogue
{
    /// <summary>
    /// Convertit un <see cref="ModeleAppareil"/> du catalogue local en <see cref="AppareilIEEE"/>
    /// utilisable par <see cref="MesureOrchestrator"/>. Permet à l'orchestrator de piloter un
    /// appareil enregistré par l'utilisateur (ex: Agilent 53131A) sans qu'il soit dans l'enum
    /// historique Stanford/Racal/EIP.
    ///
    /// L'adresse GPIB provient de la détection sur le bus (pas du catalogue, qui reste portable
    /// entre postes : un même modèle peut être à des adresses différentes selon le banc).
    /// </summary>
    public static class CatalogueAdapter
    {
        /// <summary>
        /// Construit un <see cref="AppareilIEEE"/> à partir du modèle catalogue et de l'adresse
        /// détectée sur le bus. Les gates sont construites en parsant les libellés texte
        /// (ex: "100 ms", "1 s") et en injectant la valeur dans le template <c>CommandeGate</c>
        /// du modèle (ex: <c>:FREQ:ARM:STOP:TIM {0}</c>).
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
                GereSRQ = p.GereSrq,
                SRQOn = p.SrqOn ?? string.Empty,
                SRQOff = p.SrqOff ?? string.Empty,
                Gates = ConstruireGates(modele.Gates, p.CommandeGate)
            };

            return appareil;
        }

        // Échelle canonique 0..12 utilisée par l'UI (ConfigurationViewModel.GateTimes,
        // SelectionGateViewModel.GateTimes) et par ConfigAppareilsLoader._gates pour les
        // appareils legacy. Les gates du catalogue doivent se ranger dans ces mêmes slots,
        // sinon Mesure.GateIndex venant de l'UI ne tombe pas sur l'entrée correspondante.
        private static readonly double[] _secondesSlotsUi =
        {
            0.010, 0.020, 0.050, 0.100, 0.200, 0.500,
            1.0, 2.0, 5.0, 10.0, 20.0, 50.0,
            100.0, 200.0, 500.0, 1000.0 
        };

        /// <summary>
        /// Transforme la liste de libellés (ex: <c>["10 ms", "100 ms", "1 s"]</c>) en dictionnaire
        /// dont les clés sont les indices de l'échelle UI à 13 slots (0=10 ms … 12=100 s).
        /// Un libellé "1 s" du catalogue va dans le slot 6, "10 s" dans le slot 9, etc. — comme
        /// les appareils legacy. Sans cet alignement, l'index envoyé par la UI ne correspond pas
        /// aux clés du dict et <c>AppliquerGateAsync</c> ne trouve pas la gate.
        /// Si <c>CommandeGate</c> est vide, la commande reste vide (l'orchestrator sautera la
        /// programmation — cas Racal en mode Interval par ex.).
        /// </summary>
        private static Dictionary<int, GateConfig> ConstruireGates(
            List<string> libellesGates, string templateCommande)
        {
            var dict = new Dictionary<int, GateConfig>();
            foreach (var libelle in libellesGates)
            {
                double secondes = ParserGateEnSecondes(libelle);
                int slot = TrouverSlotUi(secondes);
                if (slot < 0) continue;  // Valeur non-standard, pas mappable sur l'UI.

                string commande = string.IsNullOrWhiteSpace(templateCommande)
                    ? string.Empty
                    : string.Format(CultureInfo.InvariantCulture, templateCommande, secondes);

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
        /// Parse une chaîne comme "10 ms", "1 s", "100 s" en secondes. Accepte <c>ms</c> et <c>s</c>
        /// avec ou sans espace. Retourne 1.0 par défaut si non reconnu.
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
