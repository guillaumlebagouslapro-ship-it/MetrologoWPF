using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Metrologo.Services
{
    /// <summary>
    /// Vérifie au démarrage la présence des dépendances système nécessaires aux mesures :
    ///   • Excel (COM via ProgID) — pour générer les rapports
    ///   • NI-VISA — runtime pour piloter le bus GPIB via assembly managée
    ///   • NI-488.2 (ni4882.dll) — fast-path P/Invoke pour les commandes GPIB
    ///
    /// L'absence d'une dépendance n'empêche pas le démarrage (l'utilisateur peut vouloir
    /// faire de l'administration sans Excel ni GPIB) mais on prévient via un dialog
    /// listant l'impact.
    /// </summary>
    public static class VerificationPrerequis
    {
        /// <summary>Description d'un prérequis manquant — pour affichage UI.</summary>
        public sealed class Prerequis
        {
            public string Nom { get; init; } = string.Empty;
            public string Detail { get; init; } = string.Empty;
            public string Impact { get; init; } = string.Empty;
            public string? LienTelechargement { get; init; }
            public bool Critique { get; init; }
        }

        /// <summary>
        /// Lance tous les checks et retourne la liste des prérequis manquants. Liste
        /// vide = environnement OK.
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

        // -------- Tests individuels --------

        /// <summary>Vrai si Excel est installé (clé COM <c>Excel.Application</c> présente).</summary>
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
        /// Vrai si le runtime NI-VISA est utilisable. On tente d'instancier un
        /// <c>NationalInstruments.Visa.ResourceManager</c> : ça déclenche le chargement
        /// des assemblies + DLLs natives ; si VISA n'est pas installé, la cstor jette.
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
        /// Vrai si <c>ni4882.dll</c> est trouvable + chargeable par le loader Windows.
        /// Pas d'effet de bord : on tente <c>LoadLibrary</c> et on relâche immédiatement.
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
