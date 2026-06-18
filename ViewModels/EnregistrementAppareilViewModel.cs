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
    /// VM d'enregistrement/édition d'un appareil au catalogue. Pré-rempli depuis l'IDN détecté,
    /// l'utilisateur ajuste les commandes SCPI puis sauvegarde. Formulaire fixe calqué sur le
    /// tableau métier (Impédance, Couplage, Filtre, Trigger, Modes FREQ/TIAB, Voies A/B/C).
    /// Un champ vide = option absente dans Configuration.
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
        private const string NomFiltreA    = "Filtre HF Voie A";
        private const string NomFiltreB    = "Filtre HF Voie B";
        private const string NomFiltreC    = "Filtre HF Voie C";
        private const string NomTriggerA   = "Niveau Trigger Voie A";
        private const string NomTriggerB   = "Niveau Trigger Voie B";
        private const string NomTriggerC   = "Niveau Trigger Voie C";
        private const string NomMode       = "Mode de mesure";
        private const string NomResolution = "Résolution";
        private const string NomAtomRef    = "Référence";

        // Libellés d'options (persistés dans le JSON)
        private const string Opt50Ohm   = "50 Ω";
        private const string Opt1MOhm   = "1 MΩ";
        private const string OptAC      = "AC";
        private const string OptDC      = "DC";
        private const string OptON      = "ON";
        private const string OptOFF     = "OFF";
        private const string OptFreqA   = "FREQ Voie A";
        private const string OptFreqB   = "FREQ Voie B";
        private const string OptFreqC   = "FREQ Voie C";
        private const string OptTiab    = "TIAB";
        private const string OptResAuto = "AUTO";
        private const string OptResRecip = "RECIPROCAL";
        private const string OptResCont = "CONT";
        private const string OptRefInt  = "INT";
        private const string OptRefExt  = "EXT 10MHz";

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

        /// <summary>Temps de porte standard (10 ms..1000 s) à cocher selon les capacités de l'appareil.
        /// Evite les coquilles de saisie libre ("100ms" vs "100 ms").</summary>
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

        /// <summary>Nombre de voies (1/2/3). Pilote la visibilité des sections VOIE A/B/C.
        /// Persisté sur le modèle pour rétrocompat (défaut 2).</summary>
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(ShowVoieB))]
        [NotifyPropertyChangedFor(nameof(ShowVoieC))]
        [NotifyPropertyChangedFor(nameof(IsUneVoie))]
        [NotifyPropertyChangedFor(nameof(IsDeuxVoies))]
        [NotifyPropertyChangedFor(nameof(IsTroisVoies))]
        private int _nbVoies = 2;

        /// <summary>Vrai pour piloter la visibilité de la section VOIE B.</summary>
        public bool ShowVoieB => NbVoies >= 2;

        /// <summary>Vrai pour piloter la visibilité de la section VOIE C.</summary>
        public bool ShowVoieC => NbVoies >= 3;

        // 3 booléens couplés pour les RadioButtons TwoWay (XAML ne compare pas un int sans converter).
        public bool IsUneVoie    { get => NbVoies == 1; set { if (value) NbVoies = 1; } }
        public bool IsDeuxVoies  { get => NbVoies == 2; set { if (value) NbVoies = 2; } }
        public bool IsTroisVoies { get => NbVoies == 3; set { if (value) NbVoies = 3; } }

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

        // Trigger : template avec {0} pour la valeur en volts.
        [ObservableProperty] private string _triggerA = string.Empty;
        [ObservableProperty] private string _triggerB = string.Empty;
        [ObservableProperty] private string _triggerC = string.Empty;

        // Modes de mesure
        [ObservableProperty] private string _modeFreqA = string.Empty;
        [ObservableProperty] private string _modeFreqB = string.Empty;
        [ObservableProperty] private string _modeFreqC = string.Empty;
        [ObservableProperty] private string _modeTiab  = string.Empty;

        // Résolution (mode de calcul interne). Auto=true : masqué en Configuration, choisi selon TypeMesure.
        [ObservableProperty] private string _resolutionAuto = string.Empty;
        [ObservableProperty] private string _resolutionRecip = string.Empty;
        [ObservableProperty] private string _resolutionCont = string.Empty;

        // Référence (INT / EXT 10 MHz). Auto=false : choix manuel (cas métrologie).
        [ObservableProperty] private string _refInt = string.Empty;
        [ObservableProperty] private string _refExt = string.Empty;

        // Intervalle de temps : templates SCPI propres à l'appareil.
        // Non actif = panneau masqué en Configuration, aucune commande envoyée (ex: Stanford).
        [ObservableProperty] private bool _intervalleActif;
        [ObservableProperty] private string _intervConf1Voie = string.Empty;
        [ObservableProperty] private string _intervConf2Voies = string.Empty;
        [ObservableProperty] private string _intervCmdCouplage = string.Empty;
        [ObservableProperty] private string _intervCmdImpedance = string.Empty;
        [ObservableProperty] private string _intervCmdSeuilStart = string.Empty;
        [ObservableProperty] private string _intervCmdPenteStart = string.Empty;
        [ObservableProperty] private string _intervCmdSeuilStop1 = string.Empty;
        [ObservableProperty] private string _intervCmdPenteStop1 = string.Empty;
        [ObservableProperty] private string _intervCmdHoldoff = string.Empty;

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

            // Valeurs SCPI par défaut, correctes pour la plupart des fréquencemètres modernes.
            ChaineInit = "*RST;*CLS";
            ExeMesure = ":READ?";
            CommandeGate = ":FREQ:APER {0}";

            // Tout coché par défaut : décocher les gates non supportées.
            foreach (var g in GatesOptions) g.EstCoche = true;

            // Modèles existants proposés pour cloner leurs réglages (ex. 53230A -> 53220A).
            ChargerModelesPourCopie(exclureId: null);
        }

        /// <summary>Édition d'un modèle existant (flux Admin « Gérer les appareils »).</summary>
        public EnregistrementAppareilViewModel(ModeleAppareil modeleExistant, string utilisateur)
        {
            _modeleExistant = modeleExistant;
            _utilisateurActuel = utilisateur;

            IdnDetecte = $"(enregistré le {modeleExistant.DateCreation:dd/MM/yyyy} par {modeleExistant.CreePar})";
            ChargerDepuisModele(modeleExistant, inclureIdentite: true);
            ChargerModelesPourCopie(exclureId: modeleExistant.Id);
        }

        // ---------------- Copie des réglages depuis un modèle existant ----------------

        /// <summary>Modèles du catalogue proposés pour cloner leurs réglages (53230A → 53220A...).</summary>
        public System.Collections.ObjectModel.ObservableCollection<ModeleAppareil> ModelesPourCopie { get; } = new();

        /// <summary>Modèle source sélectionné dans la liste « Copier les réglages de… ».</summary>
        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(CopierReglagesDepuisSourceCommand))]
        private ModeleAppareil? _modeleSourceCopie;

        private void ChargerModelesPourCopie(string? exclureId)
        {
            ModelesPourCopie.Clear();
            foreach (var m in CatalogueAppareilsService.Instance.Modeles
                         .Where(m => m.Id != exclureId)
                         .OrderBy(m => m.Nom))
                ModelesPourCopie.Add(m);
        }

        private bool PeutCopier() => ModeleSourceCopie != null;

        /// <summary>Recopie TOUS les réglages (SCPI, gates, voies, intervalle...) du modèle source
        /// dans le formulaire, SANS toucher à l'identité (Nom / Fabricant IDN / Modèle IDN) : on
        /// obtient une fiche distincte aux réglages clonés (ex. 53220A à partir du 53230A).</summary>
        [RelayCommand(CanExecute = nameof(PeutCopier))]
        private void CopierReglagesDepuisSource()
        {
            if (ModeleSourceCopie == null) return;
            ChargerDepuisModele(ModeleSourceCopie, inclureIdentite: false);
            JournalLog.Info(CategorieLog.Configuration, "CATALOGUE_COPIE_REGLAGES",
                $"Réglages copiés depuis « {ModeleSourceCopie.Nom} » (identité conservée).");
        }

        /// <summary>Charge les champs du formulaire depuis <paramref name="m"/>. Si
        /// <paramref name="inclureIdentite"/> est faux, on garde le Nom/Fabricant/Modèle courants
        /// (cas du clonage de réglages vers une nouvelle fiche distincte).</summary>
        private void ChargerDepuisModele(ModeleAppareil m, bool inclureIdentite)
        {
            if (inclureIdentite)
            {
                Nom = m.Nom;
                FabricantIdn = m.FabricantIdn;
                ModeleIdn = m.ModeleIdn;
            }

            // Rétrocompat : NbVoies absent des anciens modèles -> 2 par défaut.
            NbVoies = m.NbVoies > 0 ? m.NbVoies : 2;

            ChaineInit = m.Parametres.ChaineInit;
            ConfEntree = m.Parametres.ConfEntree;
            ExeMesure = m.Parametres.ExeMesure;
            CommandeGate = m.Parametres.CommandeGate;
            CommandeMesureMultiple = m.Parametres.CommandeMesureMultiple;
            CommandeFetchFresh = m.Parametres.CommandeFetchFresh;
            TermWrite = m.Parametres.TermWrite;
            TermRead = m.Parametres.TermRead;
            TailleHeader = m.Parametres.TailleHeader;
            GereSrq = m.Parametres.GereSrq;
            SrqOn = m.Parametres.SrqOn;
            SrqOff = m.Parametres.SrqOff;

            // Coche les gates du modèle (match insensible aux espaces), décoche les autres.
            var gatesConnues = new HashSet<string>(
                m.Gates.Select(g => NormaliserLibelleGate(g)),
                StringComparer.OrdinalIgnoreCase);
            foreach (var g in GatesOptions)
                g.EstCoche = gatesConnues.Contains(NormaliserLibelleGate(g.Libelle));

            EntreesTexte = string.Join(", ", m.Entrees);
            CouplagesTexte = string.Join(", ", m.Couplages);

            // On repart à blanc sur les champs de réglages avant de recharger : sinon une 2e copie
            // (ou une copie sur un formulaire déjà rempli) laisserait traîner d'anciennes valeurs.
            ReinitialiserChampsReglages();
            ChargerReglages(m.Reglages);
            ChargerIntervalle(m.Parametres.Intervalle);
        }

        /// <summary>Vide tous les champs « réglages » du formulaire (avant un rechargement/copie).</summary>
        private void ReinitialiserChampsReglages()
        {
            ImpedanceA50 = ImpedanceA1M = ImpedanceB50 = ImpedanceB1M = ImpedanceC50 = ImpedanceC1M = string.Empty;
            CouplageAAc = CouplageADc = CouplageBAc = CouplageBDc = CouplageCAc = CouplageCDc = string.Empty;
            FiltreAOn = FiltreAOff = FiltreBOn = FiltreBOff = FiltreCOn = FiltreCOff = string.Empty;
            TriggerA = TriggerB = TriggerC = string.Empty;
            ModeFreqA = ModeFreqB = ModeFreqC = ModeTiab = string.Empty;
            ResolutionAuto = ResolutionRecip = ResolutionCont = string.Empty;
            RefInt = RefExt = string.Empty;
        }

        /// <summary>Recopie les templates d'intervalle du modèle vers les champs du formulaire.</summary>
        private void ChargerIntervalle(CommandesIntervalle? iv)
        {
            if (iv == null) return;
            IntervalleActif    = iv.Actif;
            IntervConf1Voie    = iv.Conf1Voie;
            IntervConf2Voies   = iv.Conf2Voies;
            IntervCmdCouplage  = iv.Couplage;
            IntervCmdImpedance = iv.Impedance;
            IntervCmdSeuilStart = iv.SeuilStart;
            IntervCmdPenteStart = iv.PenteStart;
            IntervCmdSeuilStop1 = iv.SeuilStop1Voie;
            IntervCmdPenteStop1 = iv.PenteStop1Voie;
            IntervCmdHoldoff   = iv.Holdoff;
        }

        /// <summary>Active l'intervalle et pré-remplit avec les templates 53230A (point de départ éditable).</summary>
        [RelayCommand]
        private void RemplirIntervalleDefaut()
        {
            var d = CommandesIntervalle.Defaut53230A();
            IntervalleActif    = true;
            IntervConf1Voie    = d.Conf1Voie;
            IntervConf2Voies   = d.Conf2Voies;
            IntervCmdCouplage  = d.Couplage;
            IntervCmdImpedance = d.Impedance;
            IntervCmdSeuilStart = d.SeuilStart;
            IntervCmdPenteStart = d.PenteStart;
            IntervCmdSeuilStop1 = d.SeuilStop1Voie;
            IntervCmdPenteStop1 = d.PenteStop1Voie;
            IntervCmdHoldoff   = d.Holdoff;
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
                    m.NbVoies = NbVoies;
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
                    NbVoies = NbVoies,
                    Parametres = ConstruireParametres(),
                    Gates = GatesCochees(),
                    Entrees = SplitCSV(EntreesTexte),
                    Couplages = SplitCSV(CouplagesTexte),
                    Reglages = ConstruireReglages(),
                    DateCreation = DateTime.UtcNow,
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

        /// <summary>Construit la liste des <see cref="ReglageAppareil"/> depuis le formulaire.
        /// Un réglage est ignoré si aucune de ses options n'est renseignée (idem pour Trigger).</summary>
        private List<ReglageAppareil> ConstruireReglages()
        {
            var liste = new List<ReglageAppareil>();

            // Voie A : toujours présente.
            AjouterChoix(liste, NomImpedanceA, (Opt50Ohm, ImpedanceA50), (Opt1MOhm, ImpedanceA1M));
            AjouterChoix(liste, NomCouplageA,  (OptAC, CouplageAAc),     (OptDC, CouplageADc));
            AjouterChoix(liste, NomFiltreA,    (OptON, FiltreAOn),       (OptOFF, FiltreAOff));
            AjouterNumerique(liste, NomTriggerA, TriggerA, unite: "V");

            if (NbVoies >= 2)
            {
                AjouterChoix(liste, NomImpedanceB, (Opt50Ohm, ImpedanceB50), (Opt1MOhm, ImpedanceB1M));
                AjouterChoix(liste, NomCouplageB,  (OptAC, CouplageBAc),     (OptDC, CouplageBDc));
                AjouterChoix(liste, NomFiltreB,    (OptON, FiltreBOn),       (OptOFF, FiltreBOff));
                AjouterNumerique(liste, NomTriggerB, TriggerB, unite: "V");
            }

            // Voie C : voie HF, présente seulement sur 3 voies.
            if (NbVoies >= 3)
            {
                AjouterChoix(liste, NomImpedanceC, (Opt50Ohm, ImpedanceC50), (Opt1MOhm, ImpedanceC1M));
                AjouterChoix(liste, NomCouplageC,  (OptAC, CouplageCAc),     (OptDC, CouplageCDc));
                AjouterChoix(liste, NomFiltreC,    (OptON, FiltreCOn),       (OptOFF, FiltreCOff));
                AjouterNumerique(liste, NomTriggerC, TriggerC, unite: "V");
            }

            // Mode de mesure : AjouterChoix filtre les options vides, les champs B/C étant vides si voies masquées.
            AjouterChoix(liste, NomMode, auto: true,
                (OptFreqA, ModeFreqA),
                (OptFreqB, NbVoies >= 2 ? ModeFreqB : string.Empty),
                (OptFreqC, NbVoies >= 3 ? ModeFreqC : string.Empty),
                (OptTiab,  ModeTiab));

            // Résolution auto (CONT pour Stab, AUTO sinon). Auto=true -> invisible en Configuration.
            AjouterChoix(liste, NomResolution, auto: true,
                (OptResAuto,  ResolutionAuto),
                (OptResRecip, ResolutionRecip),
                (OptResCont,  ResolutionCont));

            // Référence : choix manuel, Auto=false.
            AjouterChoix(liste, NomAtomRef, auto: false,
                (OptRefInt, RefInt),
                (OptRefExt, RefExt));

            return liste;
        }

        private static void AjouterChoix(List<ReglageAppareil> liste, string nom, params (string libelle, string cmd)[] options)
            => AjouterChoix(liste, nom, auto: false, options);

        private static void AjouterChoix(List<ReglageAppareil> liste, string nom, bool auto, params (string libelle, string cmd)[] options)
        {
            var opts = options
                .Where(o => !string.IsNullOrWhiteSpace(o.cmd))
                .Select(o => new OptionReglage { Libelle = o.libelle, CommandeScpi = o.cmd.Trim() })
                .ToList();

            if (opts.Count == 0) return;

            liste.Add(new ReglageAppareil { Nom = nom, Type = TypeReglage.Choix, Options = opts, Auto = auto });
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
                    case NomResolution:
                        ResolutionAuto  = CmdPourOption(r, OptResAuto);
                        ResolutionRecip = CmdPourOption(r, OptResRecip);
                        ResolutionCont  = CmdPourOption(r, OptResCont);
                        break;
                    case NomAtomRef:
                        RefInt = CmdPourOption(r, OptRefInt);
                        RefExt = CmdPourOption(r, OptRefExt);
                        break;
                }
            }
        }

        private static string CmdPourOption(ReglageAppareil r, string libelleOption)
            => r.Options.FirstOrDefault(o => o.Libelle == libelleOption)?.CommandeScpi ?? string.Empty;

        // ---------------- Params IEEE ----------------

        private ParametresIeee ConstruireParametres()
        {
            // Part des paramètres existants pour préserver les champs non exposés ici
            // (Legacy, AdresseFixeParDefaut, CommandesGateParSlot, ModeRapideActif…).
            var p = _modeleExistant?.Parametres ?? new ParametresIeee();

            p.ChaineInit = ChaineInit ?? string.Empty;
            p.ConfEntree = ConfEntree ?? string.Empty;
            p.ExeMesure = ExeMesure ?? string.Empty;
            p.CommandeGate = CommandeGate ?? string.Empty;
            p.CommandeMesureMultiple = CommandeMesureMultiple ?? string.Empty;
            p.CommandeFetchFresh = CommandeFetchFresh ?? string.Empty;
            p.TermWrite = TermWrite;
            p.TermRead = TermRead;
            p.TailleHeader = TailleHeader;
            p.GereSrq = GereSrq;
            p.SrqOn = SrqOn ?? string.Empty;
            p.SrqOff = SrqOff ?? string.Empty;
            p.Intervalle = ConstruireIntervalle();

            return p;
        }

        /// <summary>Transforme les champs intervalle du formulaire en <see cref="CommandesIntervalle"/>.</summary>
        private CommandesIntervalle ConstruireIntervalle() => new()
        {
            Actif = IntervalleActif,
            Conf1Voie = IntervConf1Voie ?? string.Empty,
            Conf2Voies = IntervConf2Voies ?? string.Empty,
            Couplage = IntervCmdCouplage ?? string.Empty,
            Impedance = IntervCmdImpedance ?? string.Empty,
            SeuilStart = IntervCmdSeuilStart ?? string.Empty,
            PenteStart = IntervCmdPenteStart ?? string.Empty,
            SeuilStop1Voie = IntervCmdSeuilStop1 ?? string.Empty,
            PenteStop1Voie = IntervCmdPenteStop1 ?? string.Empty,
            Holdoff = IntervCmdHoldoff ?? string.Empty
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
