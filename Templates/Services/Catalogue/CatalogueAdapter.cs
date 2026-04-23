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

        /// <summary>
        /// Transforme la liste de libellés (ex: <c>["10 ms", "100 ms", "1 s"]</c>) en dictionnaire
        /// indexé par position, avec commande SCPI formatée via le template <c>CommandeGate</c>.
        /// Si <c>CommandeGate</c> est vide, la commande est laissée vide (l'orchestrator sautera
        /// la programmation de gate — cas Racal en mode Interval par ex.).
        /// </summary>
        private static Dictionary<int, GateConfig> ConstruireGates(
            List<string> libellesGates, string templateCommande)
        {
            var dict = new Dictionary<int, GateConfig>();
            for (int i = 0; i < libellesGates.Count; i++)
            {
                var libelle = libellesGates[i];
                double secondes = ParserGateEnSecondes(libelle);

                string commande = string.IsNullOrWhiteSpace(templateCommande)
                    ? string.Empty
                    : string.Format(CultureInfo.InvariantCulture, templateCommande, secondes);

                dict[i] = new GateConfig
                {
                    Libelle = libelle,
                    Commande = commande,
                    ValeurSecondes = secondes
                };
            }
            return dict;
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
