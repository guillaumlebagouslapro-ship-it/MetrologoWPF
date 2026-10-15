using System.Runtime.InteropServices;

namespace Metrologo.Services
{
    /// <summary>
    /// Contrôle au démarrage des dépendances système des mesures :
    ///   • Excel (COM via ProgID) — rapports
    ///   • NI-VISA — pilotage du bus GPIB via l'assembly managée
    ///   • NI-488.2 (ni4882.dll) — fast-path P/Invoke pour les commandes GPIB
    /// Un prérequis manquant n'est pas bloquant (l'administration reste possible sans Excel ni
    /// GPIB) : un dialogue en détaille simplement l'impact.
    /// </summary>
    public static class VerificationPrerequis
    {
        /// <summary>Prérequis manquant, tel qu'affiché à l'utilisateur.</summary>
        public sealed class Prerequis
        {
            public string Nom { get; init; } = string.Empty;
            public string Detail { get; init; } = string.Empty;
            public string Impact { get; init; } = string.Empty;
            public string? LienTelechargement { get; init; }
            public bool Critique { get; init; }
        }

        /// <summary>
        /// Renvoie la liste des prérequis manquants (vide : environnement correct).
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

        /// <summary>Excel est-il installé (ProgID COM <c>Excel.Application</c>) ?</summary>
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
        /// Le runtime NI-VISA est-il exploitable ? Instancier un <c>NationalInstruments.Visa.ResourceManager</c>
        /// force le chargement des assemblies et DLLs natives ; le constructeur lève si VISA est absent.
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
        /// <c>ni4882.dll</c> est-elle chargeable ? (<c>LoadLibrary</c> aussitôt relâché.)
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
