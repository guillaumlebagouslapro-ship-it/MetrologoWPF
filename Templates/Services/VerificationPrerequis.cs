using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Metrologo.Services
{
    /// <summary>
    /// Au démarrage, on s'assure que les dépendances système dont les mesures ont besoin sont
    /// bien là :
    ///   • Excel (COM via ProgID) — pour générer les rapports
    ///   • NI-VISA — le runtime qui pilote le bus GPIB via l'assembly managée
    ///   • NI-488.2 (ni4882.dll) — le fast-path P/Invoke pour les commandes GPIB
    ///
    /// Il manque l'une d'elles ? Ce n'est pas bloquant : on peut très bien vouloir faire de
    /// l'administration sans Excel ni GPIB. On se contente d'avertir via un dialog qui détaille
    /// l'impact.
    /// </summary>
    public static class VerificationPrerequis
    {
        /// <summary>Décrit un prérequis manquant, tel qu'on l'affiche à l'utilisateur.</summary>
        public sealed class Prerequis
        {
            public string Nom { get; init; } = string.Empty;
            public string Detail { get; init; } = string.Empty;
            public string Impact { get; init; } = string.Empty;
            public string? LienTelechargement { get; init; }
            public bool Critique { get; init; }
        }

        /// <summary>
        /// Passe tous les contrôles en revue et renvoie la liste de ce qui manque. Une liste
        /// vide veut donc dire que l'environnement est bon.
        /// </summary>
        public static List<Prerequis> VerifierTout()
        {
            var manquants = new List<Prerequis>();

            if (!ExcelInstalle())
            {
                manquants.Add(new Prerequis
                {
                    Nom = "Microsoft Excel",
                    Detail = "Excel n'est pas détecté sur ce poste (clé COM Excel.Application absente).",
                    Impact = "Les mesures et la génération des rapports Excel ne fonctionneront pas. "
                           + "L'application peut néanmoins être utilisée pour la consultation et l'administration.",
                    LienTelechargement = null,
                    Critique = true,
                });
            }

            if (!NiVisaInstalle())
            {
                manquants.Add(new Prerequis
                {
                    Nom = "NI-VISA Runtime",
                    Detail = "Le runtime NI-VISA (National Instruments) n'est pas accessible — "
                           + "ResourceManager VISA introuvable.",
                    Impact = "Aucune communication GPIB possible. Les mesures qui pilotent un fréquencemètre, "
                           + "un générateur ou tout instrument sur le bus 488 sont impossibles.",
                    LienTelechargement = "https://www.ni.com/en/support/downloads/drivers/download.ni-visa.html",
                    Critique = true,
                });
            }

            if (!Ni488DllAccessible())
            {
                manquants.Add(new Prerequis
                {
                    Nom = "NI-488.2 (ni4882.dll)",
                    Detail = "La bibliothèque ni4882.dll n'est pas chargeable. "
                           + "Elle est livrée avec le runtime NI-488.2 (souvent installé en même temps que NI-VISA).",
                    Impact = "Les mesures GPIB rapides (fast-path P/Invoke) ne sont pas disponibles. "
                           + "Si NI-VISA est présent, l'app retombera sur la voie managée — un peu plus lente.",
                    LienTelechargement = "https://www.ni.com/en/support/downloads/drivers/download.ni-488-2.html",
                    Critique = false,
                });
            }

            return manquants;
        }

        // -------- Les tests pris un par un --------

        /// <summary>Renvoie vrai si Excel est installé, repéré via la clé COM <c>Excel.Application</c>.</summary>
        public static bool ExcelInstalle()
        {
            try
            {
                return Type.GetTypeFromProgID("Excel.Application") != null;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Renvoie vrai si le runtime NI-VISA est exploitable. L'astuce : on essaie d'instancier
        /// un <c>NationalInstruments.Visa.ResourceManager</c>, ce qui force le chargement des
        /// assemblies et des DLLs natives. Si VISA n'est pas installé, le constructeur lève.
        /// </summary>
        public static bool NiVisaInstalle()
        {
            try
            {
                using var rm = new NationalInstruments.Visa.ResourceManager();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Renvoie vrai si le loader Windows arrive à trouver et charger <c>ni4882.dll</c>.
        /// Sans effet de bord : on fait un <c>LoadLibrary</c> et on relâche aussitôt.
        /// </summary>
        public static bool Ni488DllAccessible()
        {
            IntPtr h = LoadLibrary("ni4882.dll");
            if (h == IntPtr.Zero) return false;
            FreeLibrary(h);
            return true;
        }

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool FreeLibrary(IntPtr hModule);
    }
}
