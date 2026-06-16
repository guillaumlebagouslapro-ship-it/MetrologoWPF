using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Metrologo.Models;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

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

        private static readonly JsonSerializerOptions _jsonOpts = new()
        {
            WriteIndented = true,
            PropertyNamingPolicy = null   // garde la casse C#, JSON facile à éditer à la main
        };

        /// <summary>
        /// Charge les profils legacy depuis le fichier réseau
        /// (<see cref="CheminsMetrologo.FichierAppareilsLegacy"/>) et les ajoute au catalogue
        /// (idempotent). Au 1er lancement le fichier n'existe pas : on l'écrit avec les valeurs
        /// par défaut ci-dessous, puis on le relit aux lancements suivants — l'utilisateur peut
        /// donc corriger les commandes GPIB dans le JSON sans recompiler.
        /// Si le réseau est injoignable, on retombe sur les profils par défaut en mémoire.
        /// </summary>
        public static void EnsureSeeded()
        {
            var defauts = new List<ModeleAppareil> { Stanford(), Racal(), Eip() };
            var profils = ChargerOuCreerFichier(defauts);

            // Garde-fou : appareils-legacy.json est éditable à la main et peut être périmé ou
            // incomplet (créé par une ancienne version, ou sauvegardé alors qu'un profil était
            // cassé). Un legacy auquel il manque un champ critique (ExeMesure, gates, init, ou
            // l'attente SRQ pour ceux qui la gèrent) envoie sa config mais NE LIT RIEN à la mesure.
            // On valide donc chaque profil chargé et on retombe sur le profil de référence en dur
            // s'il est invalide — le fichier ne peut plus neutraliser silencieusement la mesure.
            var parId = new Dictionary<string, ModeleAppareil>();
            foreach (var d in defauts) parId[d.Id] = d;

            for (int i = 0; i < profils.Count; i++)
            {
                var m = profils[i];
                if (m == null) continue;
                if (!EstProfilLegacyValide(m, parId.TryGetValue(m.Id, out var def) ? def : null)
                    && def != null)
                {
                    JournalLog.Warn(CategorieLog.Configuration, "APPAREILS_LEGACY_INCOMPLET",
                        $"Profil legacy « {m.Nom} » (Id {m.Id}) incomplet/périmé dans appareils-legacy.json "
                      + "(ExeMesure / gates / init / SRQ manquant) — profil de référence en dur réappliqué.");
                    profils[i] = def;
                }
            }
            // Réinjecte un profil de référence absent du fichier (fichier partiel).
            foreach (var d in defauts)
                if (!profils.Exists(p => p != null && p.Id == d.Id)) profils.Add(d);

            // Remplace (et non « ajoute si absent ») : si une copie corrompue d'un legacy a fui
            // dans le catalogue principal (appareils.json), on la réécrase ici par le profil de
            // référence — auto-réparation à chaque démarrage.
            var svc = CatalogueAppareilsService.Instance;
            foreach (var m in profils)
                svc.RemplacerOuAjouterEnMemoire(m);
        }

        /// <summary>
        /// Vrai si un profil legacy chargé depuis le fichier est exploitable pour mesurer : champs
        /// de communication critiques présents. Si <paramref name="reference"/> (profil en dur de
        /// même Id) est fourni et qu'il gère le SRQ, on exige aussi que le profil chargé le gère
        /// (sinon la lecture part avant le MAV et ne récupère rien — cas typique de l'EIP / Racal).
        /// </summary>
        private static bool EstProfilLegacyValide(ModeleAppareil m, ModeleAppareil? reference)
        {
            var p = m.Parametres;
            if (p == null) return false;
            if (!p.Legacy) return false;
            if (string.IsNullOrWhiteSpace(p.ExeMesure)) return false;
            if (string.IsNullOrWhiteSpace(p.ChaineInit)) return false;
            if (p.CommandesGateParSlot == null || p.CommandesGateParSlot.Count == 0) return false;

            // Cohérence SRQ : un appareil qui DOIT attendre le MAV (référence GereSrq=true) mais
            // dont le profil chargé a perdu GereSrq / SrqOn / SrqOff lirait trop tôt → valeur vide.
            if (reference?.Parametres?.GereSrq == true)
            {
                if (!p.GereSrq) return false;
                if (string.IsNullOrWhiteSpace(p.SrqOn) || string.IsNullOrWhiteSpace(p.SrqOff)) return false;
            }
            return true;
        }

        /// <summary>
        /// Réécrit le fichier réseau des profils legacy avec les modèles legacy actuellement en
        /// catalogue (notamment leur <c>AdresseFixeParDefaut</c> éditée en Administration).
        /// Réservé à l'admin. Retourne le chemin écrit, ou lève si le réseau est injoignable.
        /// </summary>
        public static string Sauvegarder()
        {
            var legacy = CatalogueAppareilsService.Instance.Modeles
                .Where(m => m.Parametres.Legacy)
                .ToList();

            string fichier = CheminsMetrologo.FichierAppareilsLegacy;
            Directory.CreateDirectory(Path.GetDirectoryName(fichier)!);
            File.WriteAllText(fichier, JsonSerializer.Serialize(legacy, _jsonOpts));

            JournalLog.Info(CategorieLog.Administration, "APPAREILS_LEGACY_SAVE",
                $"{legacy.Count} profil(s) legacy sauvegardé(s) dans {fichier}.");
            return fichier;
        }

        private static List<ModeleAppareil> ChargerOuCreerFichier(List<ModeleAppareil> defauts)
        {
            string fichier = CheminsMetrologo.FichierAppareilsLegacy;
            try
            {
                if (File.Exists(fichier))
                {
                    string json = File.ReadAllText(fichier);
                    var charges = JsonSerializer.Deserialize<List<ModeleAppareil>>(json, _jsonOpts);
                    if (charges != null && charges.Count > 0)
                    {
                        JournalLog.Info(CategorieLog.Configuration, "APPAREILS_LEGACY_LOAD",
                            $"{charges.Count} profil(s) legacy lus depuis {fichier}.");
                        return charges;
                    }
                    JournalLog.Warn(CategorieLog.Configuration, "APPAREILS_LEGACY_VIDE",
                        $"Fichier legacy vide/illisible ({fichier}) — profils par défaut utilisés.");
                    return defauts;
                }

                // 1er lancement : on écrit les profils par défaut pour que l'utilisateur les édite.
                Directory.CreateDirectory(Path.GetDirectoryName(fichier)!);
                File.WriteAllText(fichier, JsonSerializer.Serialize(defauts, _jsonOpts));
                JournalLog.Info(CategorieLog.Configuration, "APPAREILS_LEGACY_CREATE",
                    $"Fichier profils legacy créé : {fichier}. Éditez-le pour corriger les commandes GPIB.");
                return defauts;
            }
            catch (Exception ex)
            {
                JournalLog.Warn(CategorieLog.Configuration, "APPAREILS_LEGACY_ERR",
                    $"Fichier legacy inaccessible ({fichier}) : {ex.Message} — profils par défaut en mémoire.");
                return defauts;
            }
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
            NbVoies = 2,
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
                // Pas de Voie C physique sur le SR620 : comme dans le Delphi
                // (F_ConfigStanford.pas, AS_INPUT = term1,0 / term1,1 / term1,2), l'entrée est
                // un sélecteur unique 50 Ω / 1 MΩ / UHF — le passage en UHF se fait par
                // « term1,2 », pas par une voie séparée. Dans le legacy le couplage est masqué
                // quand l'entrée est UHF (le SR620 l'ignore dans ce mode).
                Choix("Entrée Voie A", ("50 Ω", "term1,0"), ("1 MΩ", "term1,1"), ("UHF", "term1,2")),
                Choix("Couplage Voie A", ("AC", "tcpl1,1"), ("DC", "tcpl1,0")),
                // Voie B (canal 2) : memes parametres, l'indice de canal change dans la
                // commande (term1/tcpl1 -> term2/tcpl2), UHF inclus. A confirmer au manuel SR620.
                Choix("Entrée Voie B", ("50 Ω", "term2,0"), ("1 MΩ", "term2,1"), ("UHF", "term2,2")),
                Choix("Couplage Voie B", ("AC", "tcpl2,1"), ("DC", "tcpl2,0")),
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
            // EIP 545 : pas de couplage / impédance / filtre / trigger — seule la bande est
            // réglable. Sélecteur identique au Delphi (F_ConfigEIP.pas, AS_INPUT = B1/B2/B3,
            // rgpGammesF « Gammes de fréquence »). Les plages viennent de la doc constructeur
            // EIP 545A (les libellés du .dfm Delphi n'ont pas été conservés) :
            //   Bande 1 : 10 Hz – 100 MHz (1 MΩ, BNC)
            //   Bande 2 : 10 MHz – 1 GHz  (50 Ω, BNC)
            //   Bande 3 : 1 – 18 GHz      (50 Ω, type N)
            // L'opérateur choisit la bande correspondant à la fréquence à mesurer ;
            // ConfEntree reste « B1 » à l'init, identique au legacy.
            Reglages = new List<ReglageAppareil>
            {
                Choix("Bande de fréquence",
                    ("Bande 1 (10 Hz – 100 MHz)", "B1"),
                    ("Bande 2 (10 MHz – 1 GHz)", "B2"),
                    ("Bande 3 (1 – 18 GHz)", "B3")),
            }
        };
    }
}
