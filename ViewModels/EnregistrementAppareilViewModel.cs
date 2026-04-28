using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Metrologo.Models;
using Metrologo.Services.Catalogue;
using Metrologo.Services.Journal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using JournalLog = Metrologo.Services.Journal.Journal;

namespace Metrologo.ViewModels
{
    /// <summary>
    /// VM de la fenêtre d'enregistrement d'un nouvel appareil au catalogue local.
    /// Pré-rempli à partir de l'IDN détecté, l'utilisateur ajuste les commandes SCPI puis sauvegarde.
    ///
    /// Les réglages exposés dans la fenêtre Configuration sont saisis via un formulaire fixe
    /// calqué sur le tableau métier (Impédance, Couplage, Filtre, Trigger, Modes FREQ/TIAB
    /// sur Voie A et Voie B). Les champs vides sont ignorés : l'option/le réglage correspondant
    /// n'apparaîtra pas dans Configuration.
    /// </summary>
    public partial class EnregistrementAppareilViewModel : ObservableObject
    {
        // ---------------- Noms canoniques des réglages (utilisés pour sérialiser et recharger) ----------------

        private const string NomImpedanceA = "Impédance Voie A";
        private const string NomImpedanceB = "Impédance Voie B";
        private const string NomImpedanceC = "Impédance Voie C";
        private const string NomCouplageA  = "Couplage Voie A";
        private const string NomCouplageB  = "Couplage Voie B";
        private const string NomCouplageC  = "Couplage Voie C";
        private const string NomFiltreA    = "Filtre Voie A";
        private const string NomFiltreB    = "Filtre Voie B";
        private const string NomFiltreC    = "Filtre Voie C";
        private const string NomTriggerA   = "Trigger Voie A";
        private const string NomTriggerB   = "Trigger Voie B";
        private const string NomTriggerC   = "Trigger Voie C";
        private const string NomMode       = "Mode de mesure";

        // Libellés d'options (persistés aussi dans le JSON)
        private const string Opt50Ohm = "50 Ω";
        private const string Opt1MOhm = "1 MΩ";
        private const string OptAC    = "AC";
        private const string OptDC    = "DC";
        private const string OptON    = "ON";
        private const string OptOFF   = "OFF";
        private const string OptFreqA = "FREQ Voie A";
        private const string OptFreqB = "FREQ Voie B";
        private const string OptFreqC = "FREQ Voie C";
        private const string OptTiab  = "TIAB";

        // ---------------- Champs de la fenêtre ----------------

        private readonly AppareilDetecte? _detecte;
        private readonly ModeleAppareil? _modeleExistant;
        private readonly string _utilisateurActuel;

        public string Titre => _modeleExistant != null ? "Modifier un appareil" : "Enregistrer un appareil";
        public bool EstModification => _modeleExistant != null;

        [ObservableProperty] private string _nom = string.Empty;
        [ObservableProperty] private string _fabricantIdn = string.Empty;
        [ObservableProperty] private string _modeleIdn = string.Empty;

        [ObservableProperty] private string _chaineInit = "*RST;*CLS";
        [ObservableProperty] private string _confEntree = string.Empty;
        [ObservableProperty] private string _exeMesure = ":READ?";
        [ObservableProperty] private string _commandeGate = ":FREQ:APER {0}";
        [ObservableProperty] private string _commandeMesureMultiple = string.Empty;
        [ObservableProperty] private string _commandeFetchFresh = string.Empty;

        [ObservableProperty] private int _termWrite = 1;
        [ObservableProperty] private int _termRead = 10;
        [ObservableProperty] private int _tailleHeader = 1;

        [ObservableProperty] private bool _gereSrq;
        [ObservableProperty] private string _srqOn = string.Empty;
        [ObservableProperty] private string _srqOff = string.Empty;

        /// <summary>
        /// Temps de porte proposés à l'utilisateur sous forme de cases à cocher. Liste fixe de
        /// 14 valeurs standard (de 10 ms à 1000 s) — l'utilisateur coche celles que son appareil
        /// supporte. Évite les coquilles d'une saisie libre (ex: "100ms" vs "100 ms").
        /// </summary>
        public System.Collections.ObjectModel.ObservableCollection<GateCochable> GatesOptions { get; } = new()
        {
            new("10 ms"), new("20 ms"), new("50 ms"),
            new("100 ms"), new("200 ms"), new("500 ms"),
            new("1 s"), new("2 s"), new("5 s"),
            new("10 s"), new("20 s"), new("50 s"),
            new("100 s"), new("200 s"), new("500 s"), new("1000 s")
        };

        [ObservableProperty] private string _entreesTexte = string.Empty;
        [ObservableProperty] private string _couplagesTexte = "AC, DC";

        [ObservableProperty] private string _idnDetecte = string.Empty;

        // ---------------- Formulaire Réglages : champs fixes du tableau ----------------

        // Impédance Voie A/B/C
        [ObservableProperty] private string _impedanceA50 = string.Empty;
        [ObservableProperty] private string _impedanceA1M = string.Empty;
        [ObservableProperty] private string _impedanceB50 = string.Empty;
        [ObservableProperty] private string _impedanceB1M = string.Empty;
        [ObservableProperty] private string _impedanceC50 = string.Empty;
        [ObservableProperty] private string _impedanceC1M = string.Empty;

        // Couplage Voie A/B/C
        [ObservableProperty] private string _couplageAAc = string.Empty;
        [ObservableProperty] private string _couplageADc = string.Empty;
        [ObservableProperty] private string _couplageBAc = string.Empty;
        [ObservableProperty] private string _couplageBDc = string.Empty;
        [ObservableProperty] private string _couplageCAc = string.Empty;
        [ObservableProperty] private string _couplageCDc = string.Empty;

        // Filtre Voie A/B/C
        [ObservableProperty] private string _filtreAOn = string.Empty;
        [ObservableProperty] private string _filtreAOff = string.Empty;
        [ObservableProperty] private string _filtreBOn = string.Empty;
        [ObservableProperty] private string _filtreBOff = string.Empty;
        [ObservableProperty] private string _filtreCOn = string.Empty;
        [ObservableProperty] private string _filtreCOff = string.Empty;

        // Trigger (template avec {0} pour la valeur en volts)
        [ObservableProperty] private string _triggerA = string.Empty;
        [ObservableProperty] private string _triggerB = string.Empty;
        [ObservableProperty] private string _triggerC = string.Empty;

        // Modes de mesure
        [ObservableProperty] private string _modeFreqA = string.Empty;
        [ObservableProperty] private string _modeFreqB = string.Empty;
        [ObservableProperty] private string _modeFreqC = string.Empty;
        [ObservableProperty] private string _modeTiab  = string.Empty;

        public Action<bool>? CloseAction { get; set; }
        public ModeleAppareil? Resultat { get; private set; }

        /// <summary>Création depuis un appareil fraîchement détecté (flux Diagnostic GPIB).</summary>
        public EnregistrementAppareilViewModel(AppareilDetecte detecte, string utilisateur)
        {
            _detecte = detecte;
            _utilisateurActuel = utilisateur;

            Nom = detecte.Modele ?? "Nouvel appareil";
            FabricantIdn = detecte.Fabricant ?? string.Empty;
            ModeleIdn = detecte.Modele ?? string.Empty;
            IdnDetecte = detecte.IdnBrut ?? string.Empty;

            // Valeurs SCPI standard — conviennent à la majorité des fréquencemètres modernes.
            // L'utilisateur peut les ajuster dans la section « Commandes de base » avant sauvegarde.
            ChaineInit = "*RST;*CLS";
            ExeMesure = ":READ?";
            CommandeGate = ":FREQ:APER {0}";

            // Par défaut, tous les temps de porte standards sont cochés — l'utilisateur
            // décoche ceux qui ne sont pas supportés par son appareil.
            foreach (var g in GatesOptions) g.EstCoche = true;
        }

        /// <summary>Édition d'un modèle existant (flux Admin « Gérer les appareils »).</summary>
        public EnregistrementAppareilViewModel(ModeleAppareil modeleExistant, string utilisateur)
        {
            _modeleExistant = modeleExistant;
            _utilisateurActuel = utilisateur;

            Nom = modeleExistant.Nom;
            FabricantIdn = modeleExistant.FabricantIdn;
            ModeleIdn = modeleExistant.ModeleIdn;
            IdnDetecte = $"(enregistré le {modeleExistant.DateCreation:dd/MM/yyyy} par {modeleExistant.CreePar})";

            ChaineInit = modeleExistant.Parametres.ChaineInit;
            ConfEntree = modeleExistant.Parametres.ConfEntree;
            ExeMesure = modeleExistant.Parametres.ExeMesure;
            CommandeGate = modeleExistant.Parametres.CommandeGate;
            CommandeMesureMultiple = modeleExistant.Parametres.CommandeMesureMultiple;
            CommandeFetchFresh = modeleExistant.Parametres.CommandeFetchFresh;
            TermWrite = modeleExistant.Parametres.TermWrite;
            TermRead = modeleExistant.Parametres.TermRead;
            TailleHeader = modeleExistant.Parametres.TailleHeader;
            GereSrq = modeleExistant.Parametres.GereSrq;
            SrqOn = modeleExistant.Parametres.SrqOn;
            SrqOff = modeleExistant.Parametres.SrqOff;

            // Coche les gates déjà enregistrées pour ce modèle (match insensible aux espaces).
            var gatesConnues = new HashSet<string>(
                modeleExistant.Gates.Select(g => NormaliserLibelleGate(g)),
                StringComparer.OrdinalIgnoreCase);
            foreach (var g in GatesOptions)
                g.EstCoche = gatesConnues.Contains(NormaliserLibelleGate(g.Libelle));

            EntreesTexte = string.Join(", ", modeleExistant.Entrees);
            CouplagesTexte = string.Join(", ", modeleExistant.Couplages);

            ChargerReglages(modeleExistant.Reglages);
        }

        // ---------------- Sauvegarde ----------------

        [RelayCommand]
        private async Task EnregistrerAsync()
        {
            if (string.IsNullOrWhiteSpace(Nom))
                Nom = _detecte?.Modele ?? _modeleExistant?.Nom ?? "Modèle sans nom";

            if (_modeleExistant != null)
            {
                await CatalogueAppareilsService.Instance.ModifierAsync(_modeleExistant.Id, m =>
                {
                    m.Nom = Nom.Trim();
                    m.FabricantIdn = FabricantIdn.Trim();
                    m.ModeleIdn = ModeleIdn.Trim();
                    m.Parametres = ConstruireParametres();
                    m.Gates = GatesCochees();
                    m.Entrees = SplitCSV(EntreesTexte);
                    m.Couplages = SplitCSV(CouplagesTexte);
                    m.Reglages = ConstruireReglages();
                });

                JournalLog.Info(CategorieLog.Administration, "CATALOGUE_MODIF",
                    $"Modèle « {Nom} » modifié par {_utilisateurActuel}.",
                    new { Id = _modeleExistant.Id, Nom });

                Resultat = _modeleExistant;
            }
            else
            {
                var modele = new ModeleAppareil
                {
                    Nom = Nom.Trim(),
                    FabricantIdn = FabricantIdn.Trim(),
                    ModeleIdn = ModeleIdn.Trim(),
                    Parametres = ConstruireParametres(),
                    Gates = GatesCochees(),
                    Entrees = SplitCSV(EntreesTexte),
                    Couplages = SplitCSV(CouplagesTexte),
                    Reglages = ConstruireReglages(),
                    DateCreation = DateTime.Now,
                    CreePar = _utilisateurActuel
                };

                await CatalogueAppareilsService.Instance.AjouterAsync(modele);

                JournalLog.Info(CategorieLog.Configuration, "CATALOGUE_AJOUT",
                    $"Modèle « {modele.Nom } » ajouté au catalogue par {_utilisateurActuel}.",
                    new { modele.Nom, modele.FabricantIdn, modele.ModeleIdn });

                Resultat = modele;
            }

            CloseAction?.Invoke(true);
        }

        [RelayCommand]
        private void Annuler() => CloseAction?.Invoke(false);

        // ---------------- Build / Load Réglages ----------------

        /// <summary>
        /// Transforme les champs du formulaire en <see cref="ReglageAppareil"/>. Un réglage ne
        /// sera créé que si au moins une de ses options est renseignée. Idem pour Trigger
        /// (ignoré si le template est vide).
        /// </summary>
        private List<ReglageAppareil> ConstruireReglages()
        {
            var liste = new List<ReglageAppareil>();

            AjouterChoix(liste, NomImpedanceA, (Opt50Ohm, ImpedanceA50), (Opt1MOhm, ImpedanceA1M));
            AjouterChoix(liste, NomImpedanceB, (Opt50Ohm, ImpedanceB50), (Opt1MOhm, ImpedanceB1M));
            AjouterChoix(liste, NomImpedanceC, (Opt50Ohm, ImpedanceC50), (Opt1MOhm, ImpedanceC1M));
            AjouterChoix(liste, NomCouplageA,  (OptAC, CouplageAAc),   (OptDC, CouplageADc));
            AjouterChoix(liste, NomCouplageB,  (OptAC, CouplageBAc),   (OptDC, CouplageBDc));
            AjouterChoix(liste, NomCouplageC,  (OptAC, CouplageCAc),   (OptDC, CouplageCDc));
            AjouterChoix(liste, NomFiltreA,    (OptON, FiltreAOn),     (OptOFF, FiltreAOff));
            AjouterChoix(liste, NomFiltreB,    (OptON, FiltreBOn),     (OptOFF, FiltreBOff));
            AjouterChoix(liste, NomFiltreC,    (OptON, FiltreCOn),     (OptOFF, FiltreCOff));

            AjouterNumerique(liste, NomTriggerA, TriggerA, unite: "V");
            AjouterNumerique(liste, NomTriggerB, TriggerB, unite: "V");
            AjouterNumerique(liste, NomTriggerC, TriggerC, unite: "V");

            AjouterChoix(liste, NomMode,
                (OptFreqA, ModeFreqA),
                (OptFreqB, ModeFreqB),
                (OptFreqC, ModeFreqC),
                (OptTiab,  ModeTiab));

            return liste;
        }

        private static void AjouterChoix(List<ReglageAppareil> liste, string nom, params (string libelle, string cmd)[] options)
        {
            var opts = options
                .Where(o => !string.IsNullOrWhiteSpace(o.cmd))
                .Select(o => new OptionReglage { Libelle = o.libelle, CommandeScpi = o.cmd.Trim() })
                .ToList();

            if (opts.Count == 0) return;

            liste.Add(new ReglageAppareil { Nom = nom, Type = TypeReglage.Choix, Options = opts });
        }

        private static void AjouterNumerique(List<ReglageAppareil> liste, string nom, string template, string unite)
        {
            if (string.IsNullOrWhiteSpace(template)) return;

            liste.Add(new ReglageAppareil
            {
                Nom = nom,
                Type = TypeReglage.Numerique,
                Unite = unite,
                Options = { new OptionReglage { Libelle = $"Valeur ({unite})", CommandeScpi = template.Trim() } }
            });
        }

        /// <summary>Recharge les champs du formulaire depuis la liste des réglages persistés.</summary>
        private void ChargerReglages(List<ReglageAppareil> reglages)
        {
            foreach (var r in reglages)
            {
                switch (r.Nom)
                {
                    case NomImpedanceA:
                        ImpedanceA50 = CmdPourOption(r, Opt50Ohm);
                        ImpedanceA1M = CmdPourOption(r, Opt1MOhm);
                        break;
                    case NomImpedanceB:
                        ImpedanceB50 = CmdPourOption(r, Opt50Ohm);
                        ImpedanceB1M = CmdPourOption(r, Opt1MOhm);
                        break;
                    case NomImpedanceC:
                        ImpedanceC50 = CmdPourOption(r, Opt50Ohm);
                        ImpedanceC1M = CmdPourOption(r, Opt1MOhm);
                        break;
                    case NomCouplageA:
                        CouplageAAc = CmdPourOption(r, OptAC);
                        CouplageADc = CmdPourOption(r, OptDC);
                        break;
                    case NomCouplageB:
                        CouplageBAc = CmdPourOption(r, OptAC);
                        CouplageBDc = CmdPourOption(r, OptDC);
                        break;
                    case NomCouplageC:
                        CouplageCAc = CmdPourOption(r, OptAC);
                        CouplageCDc = CmdPourOption(r, OptDC);
                        break;
                    case NomFiltreA:
                        FiltreAOn  = CmdPourOption(r, OptON);
                        FiltreAOff = CmdPourOption(r, OptOFF);
                        break;
                    case NomFiltreB:
                        FiltreBOn  = CmdPourOption(r, OptON);
                        FiltreBOff = CmdPourOption(r, OptOFF);
                        break;
                    case NomFiltreC:
                        FiltreCOn  = CmdPourOption(r, OptON);
                        FiltreCOff = CmdPourOption(r, OptOFF);
                        break;
                    case NomTriggerA:
                        TriggerA = r.Options.FirstOrDefault()?.CommandeScpi ?? string.Empty;
                        break;
                    case NomTriggerB:
                        TriggerB = r.Options.FirstOrDefault()?.CommandeScpi ?? string.Empty;
                        break;
                    case NomTriggerC:
                        TriggerC = r.Options.FirstOrDefault()?.CommandeScpi ?? string.Empty;
                        break;
                    case NomMode:
                        ModeFreqA = CmdPourOption(r, OptFreqA);
                        ModeFreqB = CmdPourOption(r, OptFreqB);
                        ModeFreqC = CmdPourOption(r, OptFreqC);
                        ModeTiab  = CmdPourOption(r, OptTiab);
                        break;
                }
            }
        }

        private static string CmdPourOption(ReglageAppareil r, string libelleOption)
            => r.Options.FirstOrDefault(o => o.Libelle == libelleOption)?.CommandeScpi ?? string.Empty;

        // ---------------- Params IEEE ----------------

        private ParametresIeee ConstruireParametres() => new()
        {
            ChaineInit = ChaineInit ?? string.Empty,
            ConfEntree = ConfEntree ?? string.Empty,
            ExeMesure = ExeMesure ?? string.Empty,
            CommandeGate = CommandeGate ?? string.Empty,
            CommandeMesureMultiple = CommandeMesureMultiple ?? string.Empty,
            CommandeFetchFresh = CommandeFetchFresh ?? string.Empty,
            TermWrite = TermWrite,
            TermRead = TermRead,
            TailleHeader = TailleHeader,
            GereSrq = GereSrq,
            SrqOn = SrqOn ?? string.Empty,
            SrqOff = SrqOff ?? string.Empty
        };

        private List<string> GatesCochees()
        {
            var liste = new List<string>();
            foreach (var g in GatesOptions)
            {
                if (g.EstCoche) liste.Add(g.Libelle);
            }
            return liste;
        }

        /// <summary>Normalise un libellé de gate pour comparer "100ms" ≈ "100 ms" ≈ "100  ms".</summary>
        private static string NormaliserLibelleGate(string libelle)
            => string.Concat(libelle.Where(c => !char.IsWhiteSpace(c))).ToLowerInvariant();

        private static List<string> SplitCSV(string texte)
        {
            var liste = new List<string>();
            if (string.IsNullOrWhiteSpace(texte)) return liste;
            foreach (var part in texte.Split(','))
            {
                var t = part.Trim();
                if (!string.IsNullOrEmpty(t)) liste.Add(t);
            }
            return liste;
        }
    }
}
