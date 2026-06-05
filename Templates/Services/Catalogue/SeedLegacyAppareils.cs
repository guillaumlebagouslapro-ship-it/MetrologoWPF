using System.Collections.Generic;
using Metrologo.Models;

namespace Metrologo.Services.Catalogue
{
    /// <summary>
    /// Seed des 3 fréquencemètres legacy (EIP 545, Racal-Dana 1996, Stanford SR620) qui ne
    /// répondent pas à <c>*IDN?</c> et ne sont donc pas détectables par le scan GPIB. Ils sont
    /// injectés en mémoire dans le catalogue au démarrage (idempotent, Ids stables) pour être
    /// sélectionnables en mode « adresses fixes ».
    ///
    /// Valeurs issues de <c>docs/Commandes-GPIB-Appareils.md</c> et <c>docs/legacy-delphi/Metrologo.ini</c>.
    /// Seed en mémoire uniquement (cf. <see cref="CatalogueAppareilsService.AjouterEnMemoireSiAbsent"/>) :
    /// pas d'écriture sur le partage réseau, fonctionne hors-ligne / en simulation.
    /// </summary>
    public static class SeedLegacyAppareils
    {
        public const string IdStanford = "legacy-stanford-sr620";
        public const string IdRacal     = "legacy-racal-dana-1996";
        public const string IdEip       = "legacy-eip-545";

        // Libellés de gate canoniques (alignés sur CatalogueAdapter._secondesSlotsUi, slots 0..12).
        private static readonly List<string> GatesCompletes = new()
        {
            "10 ms", "20 ms", "50 ms", "100 ms", "200 ms", "500 ms",
            "1 s", "2 s", "5 s", "10 s", "20 s", "50 s", "100 s"
        };

        /// <summary>Ajoute les 3 profils legacy au catalogue s'ils n'y sont pas déjà. Idempotent.</summary>
        public static void EnsureSeeded()
        {
            var svc = CatalogueAppareilsService.Instance;
            svc.AjouterEnMemoireSiAbsent(Stanford());
            svc.AjouterEnMemoireSiAbsent(Racal());
            svc.AjouterEnMemoireSiAbsent(Eip());
        }

        // ---------------- Helpers ----------------

        private static ReglageAppareil Choix(string nom, params (string lib, string cmd)[] options)
        {
            var r = new ReglageAppareil { Nom = nom, Type = TypeReglage.Choix };
            foreach (var (lib, cmd) in options)
                r.Options.Add(new OptionReglage { Libelle = lib, CommandeScpi = cmd });
            return r;
        }

        // ---------------- Stanford SR620 (adresse 16) ----------------

        private static ModeleAppareil Stanford() => new()
        {
            Id = IdStanford,
            Nom = "Stanford SR620 (legacy)",
            NbVoies = 3,
            Gates = new List<string>(GatesCompletes),
            Parametres = new ParametresIeee
            {
                ChaineInit = "*rst;*cls;mode3;autm0;levl1,0",
                ConfEntree = "term1,0;tcpl1,1",
                ExeMesure = "meas?0",
                CommandeGate = string.Empty,
                TermWrite = 2,   // EOI
                TermRead = 10,   // LF
                TailleHeader = 1,
                GereSrq = false,
                SrqOn = string.Empty,
                SrqOff = string.Empty,
                VerifArmingActive = false,
                ModeRapideActif = false,
                Legacy = true,
                AdresseFixeParDefaut = 16,
                CommandesGateParSlot = new Dictionary<int, string>
                {
                    [0] = "armm3;size1E0", [1] = "armm3;size2E0", [2] = "armm3;size5E0",
                    [3] = "armm4;size1E0", [4] = "armm4;size2E0", [5] = "armm4;size5E0",
                    [6] = "armm5;size1E0", [7] = "armm5;size2E0", [8] = "armm5;size5E0",
                    [9] = "armm5;size1E1", [10] = "armm5;size2E1", [11] = "armm5;size5E1",
                    [12] = "armm5;size1E2"
                }
            },
            Reglages = new List<ReglageAppareil>
            {
                // Voie A (canal 1) et Voie B (canal 2) : memes parametres, l'indice de canal
                // change dans la commande (term1/tcpl1 -> term2/tcpl2). A confirmer au manuel SR620
                // avant le materiel reel.
                Choix("Impédance Voie A", ("50 Ω", "term1,0"), ("1 MΩ", "term1,1")),
                Choix("Couplage Voie A", ("AC", "tcpl1,1"), ("DC", "tcpl1,0")),
                Choix("Impédance Voie B", ("50 Ω", "term2,0"), ("1 MΩ", "term2,1")),
                Choix("Couplage Voie B", ("AC", "tcpl2,1"), ("DC", "tcpl2,0")),
                Choix("Entrée Voie C", ("UHF", "term1,2")),
            }
        };

        // ---------------- Racal-Dana 1996 (adresse 15) ----------------

        private static ModeleAppareil Racal() => new()
        {
            Id = IdRacal,
            Nom = "Racal-Dana 1996 (legacy)",
            NbVoies = 2,
            Gates = new List<string>(GatesCompletes),
            Parametres = new ParametresIeee
            {
                ChaineInit = "QM0 MM1 AT0",
                ConfEntree = "FN2 AZ1 AA1",
                ExeMesure = "RE",
                CommandeGate = string.Empty,
                TermWrite = 1,   // NL
                TermRead = 10,   // LF
                TailleHeader = 3,
                GereSrq = true,
                SrqOn = "QM16",
                SrqOff = "QM0",
                VerifArmingActive = false,
                ModeRapideActif = false,
                Legacy = true,
                AdresseFixeParDefaut = 15,
                CommandesGateParSlot = new Dictionary<int, string>
                {
                    [0] = "GA1E-2", [1] = "GA2E-2", [2] = "GA5E-2",
                    [3] = "GA1E-1", [4] = "GA2E-1", [5] = "GA5E-1",
                    [6] = "GA1E0", [7] = "GA2E0", [8] = "GA5E0",
                    [9] = "GA1E1", [10] = "GA2E1", [11] = "GA5E1",
                    [12] = "GA1E2"
                }
            },
            Reglages = new List<ReglageAppareil>
            {
                Choix("Impédance Voie A", ("A 50 Ω", "FN2 AZ1"), ("A 1 MΩ", "FN2 AZ0")),
                Choix("Couplage Voie A", ("AC", "AA1"), ("DC", "AA0")),
                // Voie B = entree HF (FN1) : pas de couplage (comme dans le Delphi qui masque
                // le couplage sur cette entree).
                Choix("Entrée Voie B", ("Entrée HF", "FN1")),
            }
        };

        // ---------------- EIP 545 (adresse 16) ----------------

        private static ModeleAppareil Eip() => new()
        {
            Id = IdEip,
            Nom = "EIP 545 (legacy)",
            NbVoies = 1,
            Gates = new List<string> { "10 ms", "100 ms", "1 s" },
            Parametres = new ParametresIeee
            {
                ChaineInit = "HAOPR0FRSR00",
                ConfEntree = "B1",
                ExeMesure = "RS",
                CommandeGate = string.Empty,
                TermWrite = 1,   // NL
                TermRead = 10,   // LF
                TailleHeader = 1,
                GereSrq = true,
                SrqOn = "SR01",
                SrqOff = "SR00",
                VerifArmingActive = false,
                ModeRapideActif = false,
                Legacy = true,
                AdresseFixeParDefaut = 16,
                CommandesGateParSlot = new Dictionary<int, string>
                {
                    [0] = "R2", [3] = "R1", [6] = "R0"
                }
            },
            // EIP 545 : pas de couplage / impédance / filtre / trigger — seule la bande est réglable.
            Reglages = new List<ReglageAppareil>
            {
                Choix("Bande de fréquence", ("Bande 1", "B1"), ("Bande 2", "B2"), ("Bande 3", "B3")),
            }
        };
    }
}
