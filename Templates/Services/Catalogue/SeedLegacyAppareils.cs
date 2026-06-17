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
    /// Sème les 3 fréquencemètres legacy (EIP 545, Racal-Dana 1996, Stanford SR620). Comme ils ne
    /// répondent pas à <c>*IDN?</c>, le scan GPIB ne les voit pas : on les injecte donc en mémoire
    /// dans le catalogue au démarrage (opération idempotente, Ids stables) pour pouvoir les choisir
    /// en mode « adresses fixes ».
    ///
    /// Les valeurs viennent de <c>docs/Commandes-GPIB-Appareils.md</c> et de
    /// <c>docs/legacy-delphi/Metrologo.ini</c>. Tout se passe en mémoire
    /// (cf. <see cref="CatalogueAppareilsService.AjouterEnMemoireSiAbsent"/>) : rien n'est écrit sur
    /// le partage réseau, donc ça marche aussi hors-ligne ou en simulation.
    /// </summary>
    public static class SeedLegacyAppareils
    {
        public const string IdStanford = "legacy-stanford-sr620";
        public const string IdRacal     = "legacy-racal-dana-1996";
        public const string IdEip       = "legacy-eip-545";

        // Les libellés de gate de référence (calés sur CatalogueAdapter._secondesSlotsUi, slots 0..12).
        private static readonly List<string> GatesCompletes = new()
        {
            "10 ms", "20 ms", "50 ms", "100 ms", "200 ms", "500 ms",
            "1 s", "2 s", "5 s", "10 s", "20 s", "50 s", "100 s"
        };

        private static readonly JsonSerializerOptions _jsonOpts = new()
        {
            WriteIndented = true,
            PropertyNamingPolicy = null   // on garde la casse C#, le JSON reste facile à éditer à la main
        };

        /// <summary>
        /// Charge les profils legacy depuis le fichier réseau
        /// (<see cref="CheminsMetrologo.FichierAppareilsLegacy"/>) et les verse dans le catalogue
        /// (de façon idempotente). Au tout premier lancement le fichier n'existe pas encore : on
        /// l'écrit avec les valeurs par défaut ci-dessous, et on le relit aux lancements suivants —
        /// du coup l'utilisateur peut corriger les commandes GPIB directement dans le JSON, sans
        /// recompiler. Si le réseau est injoignable, on se rabat sur les profils par défaut en mémoire.
        /// </summary>
        public static void EnsureSeeded()
        {
            var defauts = new List<ModeleAppareil> { Stanford(), Racal(), Eip() };
            var profils = ChargerOuCreerFichier(defauts);

            // Garde-fou : appareils-legacy.json s'édite à la main, donc il peut très bien être
            // périmé ou incomplet (écrit par une vieille version, ou sauvegardé pendant qu'un profil
            // était cassé). Or un legacy qui a perdu un champ critique (ExeMesure, gates, init, ou
            // l'attente SRQ pour ceux qui en ont besoin) envoie bien sa config mais NE LIT RIEN à la
            // mesure. Du coup on valide chaque profil chargé et, s'il est invalide, on revient au
            // profil de référence codé en dur — comme ça le fichier ne peut plus saboter la mesure
            // dans notre dos.
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
            // On réinjecte tout profil de référence qui manquerait au fichier (cas d'un fichier partiel).
            foreach (var d in defauts)
                if (!profils.Exists(p => p != null && p.Id == d.Id)) profils.Add(d);

            // On remplace (plutôt que « ajouter si absent ») : si une copie corrompue d'un legacy
            // a fini par fuiter dans le catalogue principal (appareils.json), on l'écrase ici par
            // le profil de référence — une petite auto-réparation à chaque démarrage.
            var svc = CatalogueAppareilsService.Instance;
            foreach (var m in profils)
                svc.RemplacerOuAjouterEnMemoire(m);
        }

        /// <summary>
        /// Renvoie vrai si un profil legacy lu depuis le fichier est réellement utilisable pour
        /// mesurer, c'est-à-dire si ses champs de communication critiques sont bien là. Quand on
        /// passe une <paramref name="reference"/> (le profil codé en dur de même Id) et qu'elle
        /// gère le SRQ, on exige en plus que le profil chargé le gère aussi — sans ça la lecture
        /// démarre avant le MAV et ne ramène rien (le grand classique de l'EIP / Racal).
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
            // dont le profil chargé aurait perdu GereSrq / SrqOn / SrqOff lirait trop tôt → on
            // récupérerait une valeur vide.
            if (reference?.Parametres?.GereSrq == true)
            {
                if (!p.GereSrq) return false;
                if (string.IsNullOrWhiteSpace(p.SrqOn) || string.IsNullOrWhiteSpace(p.SrqOff)) return false;
            }
            return true;
        }

        /// <summary>
        /// Réécrit le fichier réseau des profils legacy à partir des modèles legacy actuellement
        /// dans le catalogue (avec, entre autres, leur <c>AdresseFixeParDefaut</c> modifiée en
        /// Administration). Réservé à l'admin. Renvoie le chemin écrit, ou lève si le réseau est
        /// injoignable.
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

                // Premier lancement : on dépose les profils par défaut pour que l'utilisateur puisse les éditer.
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
                // Le SR620 n'a pas de Voie C physique : comme dans le Delphi
                // (F_ConfigStanford.pas, AS_INPUT = term1,0 / term1,1 / term1,2), l'entrée n'est
                // qu'un sélecteur unique 50 Ω / 1 MΩ / UHF — on bascule en UHF avec « term1,2 »,
                // pas via une voie à part. Dans le legacy, le couplage est masqué dès que l'entrée
                // passe en UHF (le SR620 n'en tient pas compte dans ce mode).
                Choix("Entrée Voie A", ("50 Ω", "term1,0"), ("1 MΩ", "term1,1"), ("UHF", "term1,2")),
                Choix("Couplage Voie A", ("AC", "tcpl1,1"), ("DC", "tcpl1,0")),
                // Voie B (canal 2) : memes parametres, seul l'indice de canal change dans la
                // commande (term1/tcpl1 -> term2/tcpl2), UHF compris. A confirmer dans le manuel SR620.
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
                // Voie B = l'entree HF (FN1) : pas de couplage (le Delphi masquait deja le
                // couplage sur cette entree).
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
            // EIP 545 : ni couplage, ni impédance, ni filtre, ni trigger — il n'y a que la bande
            // à régler. Le sélecteur est calqué sur le Delphi (F_ConfigEIP.pas, AS_INPUT = B1/B2/B3,
            // rgpGammesF « Gammes de fréquence »). Les plages proviennent de la doc constructeur
            // EIP 545A (on n'a pas gardé les libellés du .dfm Delphi) :
            //   Bande 1 : 10 Hz – 100 MHz (1 MΩ, BNC)
            //   Bande 2 : 10 MHz – 1 GHz  (50 Ω, BNC)
            //   Bande 3 : 1 – 18 GHz      (50 Ω, type N)
            // L'opérateur prend la bande qui correspond à la fréquence qu'il veut mesurer ;
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
