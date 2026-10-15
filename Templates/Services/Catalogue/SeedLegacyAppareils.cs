using System.IO;
using System.Text.Json;
using Metrologo.Models;
using Metrologo.Services.Journal;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.Services.Catalogue
{
    /// <summary>
    /// Sème les 3 fréquencemètres legacy (EIP 545, Racal-Dana 1996, Stanford SR620). Ils ne répondent
    /// pas à <c>*IDN?</c> (le scan GPIB ne les voit pas) : ils sont injectés en mémoire dans le
    /// catalogue à chaque démarrage (idempotent, Ids stables) pour le mode « adresses fixes ».
    /// Valeurs issues de <c>docs/Commandes-GPIB-Appareils.md</c> et <c>docs/legacy-delphi/Metrologo.ini</c>.
    /// Rien n'est écrit dans le catalogue partagé (fonctionne hors-ligne et en simulation).
    /// </summary>
    public static class SeedLegacyAppareils
    {
        public const string IdStanford = "legacy-stanford-sr620";
        public const string IdRacal     = "legacy-racal-dana-1996";
        public const string IdEip       = "legacy-eip-545";

        // Gates de référence (13 premiers slots de CatalogueAdapter._secondesSlotsUi).
        private static readonly List<string> GatesCompletes = new()
        {
            "10 ms", "20 ms", "50 ms", "100 ms", "200 ms", "500 ms",
            "1 s", "2 s", "5 s", "10 s", "20 s", "50 s", "100 s"
        };

        private static readonly JsonSerializerOptions _jsonOpts = new()
        {
            WriteIndented = true,
            PropertyNamingPolicy = null   // casse C# conservée (JSON édité à la main)
        };

        /// <summary>
        /// Charge les profils legacy depuis <see cref="CheminsMetrologo.FichierAppareilsLegacy"/> et
        /// les verse (idempotent) dans le catalogue. Le fichier est créé avec les valeurs par défaut
        /// ci-dessous au 1er lancement puis relu : les commandes GPIB se corrigent dans le JSON sans
        /// recompiler. Réseau injoignable : profils par défaut en mémoire.
        /// </summary>
        public static void EnsureSeeded()
        {
            var defauts = new List<ModeleAppareil> { Stanford(), Racal(), Eip() };
            var profils = ChargerOuCreerFichier(defauts);

            // Garde-fou : appareils-legacy.json est édité à la main et peut être périmé ou incomplet.
            // Un legacy sans champ critique (ExeMesure, gates, init, SRQ) envoie sa config mais ne lit
            // rien à la mesure : tout profil invalide est remplacé par le profil de référence codé en dur.
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
            // Profils de référence absents du fichier (fichier partiel) réinjectés.
            foreach (var d in defauts)
                if (!profils.Exists(p => p != null && p.Id == d.Id)) profils.Add(d);

            // Remplacement (et non « ajouter si absent ») : écrase une copie corrompue d'un legacy
            // qui aurait fuité dans appareils.json (auto-réparation à chaque démarrage).
            var svc = CatalogueAppareilsService.Instance;
            foreach (var m in profils)
                svc.RemplacerOuAjouterEnMemoire(m);
        }

        /// <summary>
        /// Le profil legacy lu est-il utilisable pour mesurer (champs de communication critiques
        /// présents) ? Si la <paramref name="reference"/> (profil codé en dur de même Id) gère le SRQ,
        /// le profil chargé doit le gérer aussi : sinon la lecture démarre avant le MAV et ne ramène
        /// rien (cas classique de l'EIP / Racal).
        /// </summary>
        private static bool EstProfilLegacyValide(ModeleAppareil m, ModeleAppareil? reference)
        {
            var p = m.Parametres;
            if (p == null) return false;
            if (!p.Legacy) return false;
            if (string.IsNullOrWhiteSpace(p.ExeMesure)) return false;
            if (string.IsNullOrWhiteSpace(p.ChaineInit)) return false;
            if (p.CommandesGateParSlot == null || p.CommandesGateParSlot.Count == 0) return false;

            if (reference?.Parametres?.GereSrq == true)
            {
                if (!p.GereSrq) return false;
                if (string.IsNullOrWhiteSpace(p.SrqOn) || string.IsNullOrWhiteSpace(p.SrqOff)) return false;
            }
            return true;
        }

        /// <summary>
        /// Réécrit le fichier réseau des profils legacy à partir du catalogue (dont
        /// <c>AdresseFixeParDefaut</c> modifiée en Administration). Réservé à l'admin. Renvoie le
        /// chemin écrit ; lève si le réseau est injoignable.
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

                // Premier lancement : dépôt des profils par défaut, éditables.
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
                // Le SR620 n'a pas de Voie C physique : comme dans le Delphi (F_ConfigStanford.pas,
                // AS_INPUT = term1,0 / term1,1 / term1,2), l'entrée est un sélecteur unique
                // 50 Ω / 1 MΩ / UHF (« term1,2 »). Le legacy masque le couplage en UHF (ignoré par le SR620).
                Choix("Entrée Voie A", ("50 Ω", "term1,0"), ("1 MΩ", "term1,1"), ("UHF", "term1,2")),
                Choix("Couplage Voie A", ("AC", "tcpl1,1"), ("DC", "tcpl1,0")),
                // Voie B (canal 2) : seul l'indice de canal change (term1/tcpl1 -> term2/tcpl2),
                // UHF compris. TODO : à confirmer dans le manuel SR620.
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
                // Voie B = entrée HF (FN1), sans couplage (déjà masqué dans le Delphi).
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
            // EIP 545 : seule la bande se règle (ni couplage, impédance, filtre, trigger). Sélecteur
            // calqué sur le Delphi (F_ConfigEIP.pas, AS_INPUT = B1/B2/B3) ; plages issues de la doc
            // constructeur EIP 545A :
            //   Bande 1 : 10 Hz – 100 MHz (1 MΩ, BNC)
            //   Bande 2 : 10 MHz – 1 GHz  (50 Ω, BNC)
            //   Bande 3 : 1 – 18 GHz      (50 Ω, type N)
            // ConfEntree reste « B1 » à l'init, comme dans le legacy.
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
