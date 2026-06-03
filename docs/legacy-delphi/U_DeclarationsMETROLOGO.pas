Unit U_DeclarationsMETROLOGO;

Interface

Uses
    Windows, Classes, U_OutilsIEEE, dpib32, U_OutilsString, uE2M, SysUtils, Dialogs,
    Forms, JvSearchFiles, Variants, F_OutilsSQL, ExtCtrls, U_OutilsSystemes,
    Registry;

Type
// --- Enumération des modes de fonctionnement de Metrologo ---
    emModesMetrologo = (emExploitation = 0, emSimulation, emValidation);

// --- Enumeration des fréquencemètres ---
    enAppareilsIEEE = (eaStanford = 0, eaRacal, eaEIP);

// --- Enumeration des types de mesures possibles ---
    enTypesMesures = (etFreqAvantInterv = 0, etStabilite, etDerive, etFrequence, etFreqFinale, etInterval, etTachyContact, etTachyOptique, etStroboscope);
//    TTypeMesure = Set Of enTypesMesures;

// --- Enumeration des modes de mesure ---
    enModeMesures = (emDirect, emIndirect);

// --- Enumeration des configuration possibles pour les entrées de chaque fréquencemetre ---
    enInputStanford = (eiStanfordLowZ = 0, eiStanford1MOhm, eiStanfordUHF);
    enInputRacal = (eiInputA50 = 0, eiInputA1M, eiInputC);
    enInputEIP = (eiBand1 = 0, eiBand2, eiBand3);

// --- Enumération des coupling d'entrée des fréquencemètres ---
    enInputCoupling = (eiAC = 0, eiDC);

// --- Zones nommées du classeur de validation ---
    enZNValidation =
        (evCorrFreq = 0,
        evValMoy,
        evETIncertGlob,
        evMultFreq,
        evFreqLow,
        evFreqHigh,
        evIncertResol,
        evIncertSupp,
        evIncertResolEtSupp
        );

// --- Enumeration des zones nommées XL ---
    enZonesNommees =
        (I_IDX_ZNNOFICHE = 0,                                                   // Feuille mesure : N° de la FI.
        I_IIDX_ZNNOFICHERECAPFREQ,                                              // Feuille Récap Freq : N° de la FI.
        I_IIDX_ZNNOFICHERECAPSTAB,                                              // Feuille Récap Stab : N° de la FI.
        I_IIDX_ZNNOFICHERECAPDERIVE,                                            // Feuille Récap Dérive : N° de la FI.
        I_IDX_ZNDATE,                                                           // Feuilles mesure : Date de la mesure
        I_IDX_ZNDATERECAPFREQ,                                                  // Feuille Récap. Freq: Date de début des mesures
        I_IDX_ZNDATERECAPSTAB,                                                  // Feuille Récap. Stab: Date de début des mesures
        I_IDX_ZNDATERECAPDERIVE,                                                // Feuille Récap. Derive: Date de début des mesures
        I_IDX_ZNTYPEMESURE,                                                     // Feuille Récap. : Type de mesure (Fréquence, Stabilité, Dérive)
        I_IDX_ZNRUBIDIUM,                                                       // Feuilles mesure : Nom du rubidium utilisé pour les mesures
        I_IDX_ZNGATE,                                                           // Feuilles mesure : Temps de porte pour la mesure
        I_IDX_ZNLIBGATE,                                                        // Feuilles mesure : Libellé du temps de gate
        I_IDX_ZNMODEMESURE,                                                     // Feuilles mesure : Mode Direct/Indirect
        I_IDX_ZNCOEFFMULT,                                                      // Feuilles mesure : Coeff. multiplicateur si en mode Indirect
        I_IDX_ZNECARTTYPE,                                                      // Feuilles mesure : Valeur de l'écart-type
        I_IDX_ZNINCERTECARTTYPE,                                                // Feuilles mesure : Valeur de l'incertitide sur l'écart-type
        I_IDX_ZNFREQREF,                                                        // Feuilles mesure : Valeur de la fréquence de référence
        I_IDX_ZNFREQCORR,                                                       // Feuilles mesure : Valeur de la fréquence corrigée
        I_IDX_ZNINCERTRESOL,                                                    // Feuilles mesure : Valeur de l'incertitude sur la résolution
        I_IDX_ZNINCERTSUP,                                                      // Feuilles mesure : Valeur de l'incertitude supplémentaire
        I_IDX_ZNINCERTACCREDITEE,                                               // Feuilles mesure : Valeur de l'incertitude accréditée
        I_IDX_ZNINCERTGLOBALE,                                                  // Feuilles mesure : Valeur de l'incertitude globale
        I_IDX_ZNFREQUTILISE,                                                    // Feuilles mesure : Nom du féquencemètre utilisé
        I_IDX_ZNLIBFNOMINALE,                                                   // Feuilles mesure : Libellé pour affichage de la Freq. nominale si mesure Indirecte
        I_IDX_ZNVALFNOMINALE,                                                   // Feuilles mesure : Valeur de Freq. nominale si mesure Indirecte
        I_IDX_ZNFREQMOYREEL,                                                    // Feuilles mesure : Valeur de la fréquence moyenne réelle
        I_IDX_ZNNBMESURES,                                                      // Feuilles mesure : Nombre de mesures de la feuille
        I_IDX_ZNRECAPF_NBMES,                                                   // Feuilles Récap. fréquence : Nombre de mesures de la feuille
        I_IDX_ZNVARIANCE,                                                       // Feuilles mesure : Valeur de la variance
        I_IDX_ZNRECAPD_NBMESCYCLE,                                              // Feuille Récap. Dérive: Nombre de mesures par cycle
        I_IDX_ZNRECAPD_TPSMES,                                                  // Feuille Récap. Dérive : Temps de mesure
        I_IDX_ZNRECAPD_TPSCYCLE,                                                // Feuille Récap. Dérive : Temps du cycle de dérive
        I_IDX_ZNRECAPF_DEBZONE,                                                 // Feuilles Récap. Fréquence : Zone où coller le format de la ligne de récap. à créer
        I_IDX_ZNRECAPF_LIGNE0,                                                  // Feuilles Récap. Fréquence : Zone modèle pour création des lignes de récap Fréq.
        I_IDX_ZNRECAPS_DEBZONE,                                                 // Feuilles Récap. Stabilité : Zone où coller le format de la ligne de récap. à créer
        I_IDX_ZNRECAPS_LIGNE0,                                                  // Feuilles Récap. Stabilité : Zone modèle pour création des lignes de récap Fréq.
        I_IDX_ZNRECAPD_DEBZONE,                                                 // Feuilles Récap. Dérive : Zone où coller le format de la ligne de récap. à créer
        I_IDX_ZNRECAPD_LIGNE0,                                                  // Feuilles Récap. Dérive : Zone modèle pour création des lignes de récap Fréq.
        I_IDX_ZNCOEFFA,                                                         // Feuilles mesure : Valeur du coef A issu de la base SQL (pour calc incertitude)
        I_IDX_ZNCOEFFB,                                                         // Feuilles mesure : Valeur du coef B issu de la base SQL (pour calc incertitude)
        I_IDX_ZNVALGATESECONDES,                                                // Feuilles mesure : Valeur du temps de gate en secondes
        I_IDX_ZNNBMESACCREDITE,                                                 // Feuilles mesure : Valeur du nombre de mesures accrédité
        I_IDX_ZNTEMPSMESUREACCREDITE,                                           // Feuilles mesure : Valeur du temps de mesures accrédité
        I_IDX_ZNCOEFFNBMES,                                                     // Feuilles mesure : Coeff lié au nombre de mesures
        I_IDX_ZNCOEFFTEMPSMESURE,                                               // Feuilles mesure : Coeff lié au temps de mesure
        I_IDX_ZNINCERTGLOBNONSTAB,                                              // Feuilles mesure : Valeur de stockage du calcul de l'incert globale (Pour mesures NON Stab)
        I_IDX_ZNGRANDEVALEUR,                                                   // Feuille Récap. Stabilité : Plus grande valeur Maxi
        I_IDX_ZNPETITEVALEUR,                                                   // Feuille Récap. Stabilité : Plus petite valeur Mini
        I_IDX_ZNCOEFFDERIVEJOURNALIERE                                          // Feuiklle Recap Derive : Coeff pour calcul de la derive journaliere
        );

// --- Enumeration des valeurs de Gate ---
    enGates =
        (I_GATE10MS = 0, I_GATE20MS, I_GATE50MS, I_GATE100MS, I_GATE200MS, I_GATE500MS, I_GATE1S, I_GATE2S,
        I_GATE5S, I_GATE10S, I_GATE20S, I_GATE50S, I_GATE100S);

// --- Classe multiplexeur ---
    TMultiplexeur = Class(TObject)
        sNom: String;
        iAdresse: SmallInt;
        iWriteTerm: Integer;
        iReadTerm: Integer;
    End;

// Classe mesure
    TMesure = Class(TObject)
        TimerDerive: TTimer;

    Private
        FNumFI: String;
        FNumFISlash: String;
        FDossierResultats: String;
        FNomFicDerive: String;
        FNomFicFreq: String;
        FNomFicStab: String;
        FID_FI: Integer;
        FTimeDerive: TNotifyEvent;

        Procedure setNumFI(Const Value: String);
        Procedure TimerEllapsed(Sender: TObject);

    Public
        FNominale: Double;
        FReference: Double;
        iNbMesAccred: Integer;
        iTmpsMesAccred: Integer;
        VoieMux: Integer;
        Frequencemetre: enAppareilsIEEE;
        ModeMesure: enModeMesures;
        IndexMultiplicateur: Integer;
        TypeMesure: enTypesMesures;
//        TypeMesure: TTypeMesure;
        InputStanford: enInputStanford;
        InputRacal: enInputRacal;
        InputEIP: enInputEIP;
        CouplingStanford: enInputCoupling;
        CouplingRacal: enInputCoupling;
        IndexGate: Integer;
        IndexIntervalDerive: Integer;
        NbCyclesDeriveEffectuees: Integer;
        NbMesures: Integer;
        NbMesDerive: Integer;
        NbCyclesDerive: Integer;
        DateDebutDerive: TDateTime;
        DateNextCycleDerive: TDateTime;
        bMesureAvecFreq: Boolean;
        bInitManu: Boolean;

        Property onTimeDerive: TNotifyEvent Read FTimeDerive Write FTimeDerive;
        Property sNumFI: String Read FNumFI Write setNumFI;
        Property sNumFISlash: String Read FNumFISlash;
        Property iID_FI: Integer Read FID_FI;
        Property NomFicFreq: String Read FNomFicFreq;
        Property NomFicStab: String Read FNomFicStab;
        Property NomFicDerive: String Read FNomFicDerive;
        Property sDossierResultats: String Read FDossierResultats;

        Procedure Init();

        Procedure SaveToRegistry();
        Procedure UpdateRegistry();
        Function ReadFromRegistry(): Boolean;
        Procedure RemoveFromRegistry();

        Function IsFreqMeasurement(): Boolean;
        Function GetIndexedFile(spRacineNom: String): String;

        Constructor Create;

    End;

// --- Classe fréquencemètre ---
    TAppareilIEEE = Class(TObject)

    Public
        sNom: String;
        iAdresse: SmallInt;
        iWriteTerm: Integer;
        iReadTerm: Integer;
        iTailleHeaderReponse: Integer;
        sInit: String;
        sExeMesure: String;
        sMonocoup: String;
        sConfEntree: String;
        bGereSRQ: Boolean;
        sSRQOn: String;
        sSRQOff: String;
        lstCdeGates: TStringList;                                               // Liste des mnémonique de programmation des valeurs de gate
        lstLibellesGates: TStringList;                                          // Liste des libellés des ComboBox
        lstIdxValeursGates: TStringList;                                        // Indexs des valeurs de Gate en secondes dans tableau

//        Function Lecture(spCde: String): Double;
        Function Lecture(opMes: TMesure): Double;
        function MesureIntervalleRacalDana(ipAdresse: Smallint; ipWriteTerm, ipReadTerm: Integer): string;

    End;

    T_ZNXL = Record
        Nom: String;
        bUnique: Boolean;
    End;

    T_ZONESVALIDATION = Record
        ZNLibelle: String;
        ZNDonnees: String;
    End;

    T_InfosGates = Record
        sValGate: String;
        sLibGate: String;
        dValSecondes: Double;
    End;

    TRubidiums = Record
        ID: Integer;
        Design: String;
    End;

    TGrapheDerive = record
        iSemaine: Integer;
        iAnnee: Integer;
        dValeur: Double;
        bMoyennee: Boolean;
    end;

    TInsertGraphDerive = class(TObject)
        IndexDeb: Integer;
        DateDeb: TDateTime;
        DateFin: TDateTime;
        ValMoy: Double;
    end;

    ATDouble = Array Of Double;
    PTabDouble = ^ATDouble;

Const

    LCID: Integer = LOCALE_USER_DEFAULT;

    I_MULT_COEF0: Integer = 0;
    I_MULT_COEF1: Integer = 1;
    I_MULT_COEF2: Integer = 2;
    I_MULT_COEF3: Integer = 3;
    I_MULT_COEF4: Integer = 4;

    I_MUX_V1: Integer = 0;
    I_MUX_V2: Integer = 1;
    I_MUX_V3: Integer = 2;
    I_MUX_V4: Integer = 3;
    I_MUX_V5: Integer = 4;
    I_MUX_V6: Integer = 5;
    I_MUX_V7: Integer = 6;

    I_NBMESURES_DEFAUT: Integer = 30;                                           // Nb de mesures par défaut
    I_NBMESDERIVE_DEFAULT: Integer = 30;                                        // Nb de mesures par défaut pour la dérive

    I_IDX_PERIODICITE_DERIVE_DEFAUT: Integer = 6;                               // Index de la périodicité par défaut pour une mesure de dérive
    I_IDX_PERIODICITE_DERIVE_VALIDATION: Integer = 0;                           // Index de la périodicité pour une mesure de dérive (En mode Validation)

    I_NBCYCLESDERIVE_DEFAUT: Integer = 40;                                      // Nb cycles de la mesure de dérive par défaut
    I_NBCYCLESDERIVE_VALIDATION: Integer = 10;                                  // Nb cycles de la mesure de dérive (En mode validation)

    // Constantes indexs intervalles de dérive
    I_INTERVAL_20S = 0;
    I_INTERVAL_1MN = 1;
    I_INTERVAL_15MN = 2;
    I_INTERVAL_1H = 3;
    I_INTERVAL_3H = 4;
    I_INTERVAL_6H = 5;
    I_INTERVAL_12H = 6;
    I_INTERVAL_18H = 7;
    I_INTERVAL_1J = 8;

    // Tableau des equivalents en secondes des intervalles de cycles de dérive
    AI_INTERVAL_DERIVE: Array[I_INTERVAL_20S..I_INTERVAL_1J] Of Integer = (20, 60, 900, 3600, 10800, 21600, 43200, 64800, 86400);
    AI_LIB_INTERVAL_DERIVE: Array[I_INTERVAL_20S..I_INTERVAL_1J] of string = ('20 s', '1 min.', '15 min.', '1 h', '3 h', '6 h', '12 h', '18 h','24 h');
    // Noms des sections representant des fréquencemètres dans le fichier Metrologo.ini

    // ATTENTION l'ordre des fréquencemètres défini dans l'enum suivante DOIT être identique à celui des indexs
    // de la base SQL (Champ APP_INDEX de la table TR_METROLOGO_APPS)
    S_NOMS_APPIEE: Array[enAppareilsIEEE] Of String = ('Stanford SR620', 'Racal-Dana 1996', 'EIP 545');
    S_MUX_HP: String = 'HP59307A';

    // Noms des valeurs de chaque section 'Frequencemetre' du fichier Metrologo.ini
    S_VAL_ADRESSE: String = 'Adresse';
    S_VAL_WRITETERM: String = 'WriteTerm';
    S_VAL_READTERM: String = 'ReadTerm';
    S_VAL_HEADERREPONSE: String = 'TailleHeaderReponse';
    S_VAL_CHAINEINIT: String = 'ChaineInit';
    S_VAL_CONFENTREE: String = 'ConfEntreeDef';
    S_VAL_EXEMESURE: String = 'ExeMesure';
    S_VAL_MONOCOUP: String = 'Monocoup';
    S_VAL_GESTSRQ: String = 'GestSRQ';
    S_VAL_SRQON: String = 'SRQOn';
    S_VAL_SRQOFF: String = 'SRQOff';

    // TimeOut par défaut de la carte IEEE
    I_DEF_IEEETIMEOUT = TNONE;

    // Indexs des tableaux relatifs aux Gates

    I_IDX_PROCEDURE_EIP = -3;                                                   // Procédure auto de stab pour EIP545 (Nom permettant de rester conforme à MetroV4)
    I_IDX_PROCEDURE_10S = -2;                                                   // Procédure auto de stab (10ms à 10s) pour NON EIP545 (Nom permettant de rester conforme à MetroV4)
    I_IDX_PROCEDURE_100S = -1;                                                  // Procédure auto de stab (10ms à 100s) pour NON EIP545 (Nom permettant de rester conforme à MetroV4)

    AT_VAL_GATES: Array[enGates] Of T_InfosGates =
    (
        (sValGate: 'Gate10ms'; sLibGate: '10ms'; dValSecondes: 0.01),
        (sValGate: 'Gate20ms'; sLibGate: '20ms'; dValSecondes: 0.02),
        (sValGate: 'Gate50ms'; sLibGate: '50ms'; dValSecondes: 0.05),
        (sValGate: 'Gate100ms'; sLibGate: '100ms'; dValSecondes: 0.1),
        (sValGate: 'Gate200ms'; sLibGate: '200ms'; dValSecondes: 0.2),
        (sValGate: 'Gate500ms'; sLibGate: '500ms'; dValSecondes: 0.5),
        (sValGate: 'Gate1s'; sLibGate: '1s'; dValSecondes: 1.0),
        (sValGate: 'Gate2s'; sLibGate: '2s'; dValSecondes: 2.0),
        (sValGate: 'Gate5s'; sLibGate: '5s'; dValSecondes: 5.0),
        (sValGate: 'Gate10s'; sLibGate: '10s'; dValSecondes: 10.0),
        (sValGate: 'Gate20s'; sLibGate: '20s'; dValSecondes: 20.0),
        (sValGate: 'Gate50s'; sLibGate: '50s'; dValSecondes: 50.0),
        (sValGate: 'Gate100s'; sLibGate: '100s'; dValSecondes: 100.0)
        );

//    S_APP_CAPTION: String = 'AI D93301/A - L013 - METROLOGO  V';
    S_APP_CAPTION: String = 'L013 - METROLOGO  V';       // Demande de PB par groupwise du 29/12/2011
    S_HTML_CAPTION: String = '<P align="center">%s</P>';

    AS_CAPTION_MODEMETROLOGO: Array[emModesMetrologo] Of String = ('Mode exploitation', 'Mode simulation', 'Mode validation');

    // Informations permettant de récuperer le nom du Mutex dans T_PATH
    S_MODULENAME: String = 'MUTEXS';
    S_DENOMNAME: String = 'Metrologo';

    // Noms des champs de la base de données
    S_RUB_FMOYENNE: String = 'RUB_FMOYENNE';
    S_RUB_DESIGNATION: String = 'RUB_DESIGNATION';
    S_RUB_ID: String = 'RUB_ID';
    S_RUB_ACTIF: String = 'RUB_ACTIF';
    S_RUB_AVECGPS: String = 'RUB_AVECGPS';

    S_KEYMACRO: String = 'Software\Microsoft\Office\11.0\Excel\Options';        // Cles du registre pour chargt XLA
    S_XLA: String = 'C:\Exe_Spe\Fct_VBA\Metrologo.xla';                         // Macro complémentaire metrologo

    S_KEY_BDRMETROLOGO = 'SOFTWARE\E2M\Metrologo';
    S_KEY_BDR: String = S_KEY_BDRMETROLOGO + '\%s';

    S_KEY_REMINDERMARCHE: String = 'ReminderMarche';
    S_KEY_REMINDERFREQAUTRES: String = 'ReminderFreqAutres';
    S_KEY_COMMANDE: string ='Commande';    // LSO le 18/02/2015 Commande à envoyer sur le bus IEEE
    S_KEY_ADRESSE: string ='Adresse';    // LSO le 18/02/2015 Adresse IEEE de l'appareil

    S_FORMATDATETIME: String = 'dd/mm/yyyy hh:nn:ss';

    // Noms de clerfs de la base de registre
    S_BDR_FNOMINALE: String = 'FNominale';
    S_BDR_NBMESACCRED: String = 'NbMesAccred';
    S_BDR_TEMPSACCRED: String = 'TempsMesAccred';
    S_BDR_VOIEMUX: String = 'VoieMux';
    S_BDR_FREQ: String = 'Frequencemetre';
    S_BDR_MODEMESURE: String = 'ModeMesure';
    S_BDR_IDXMULT: String = 'Multiplicateur';
    S_BDR_TYPEMES: String = 'TypeMesure';
    S_BDR_INPUTSR620: String = 'InputSR620';
    S_BDR_INPUT1996: String = 'Input1996';
    S_BDR_INPUT545: String = 'Input545';
    S_BDR_CPLGSR620: String = 'CplgSR620';
    S_BDR_CPLG1996: String = 'Cplg1996';
    S_BDR_INDEXGATE: String = 'IndexGate';
    S_BDR_IDX_INTERVDERIVE: String = 'IdxIntervalDerive';
    S_BDR_NBMES_DERIVES_EFFECTUEES: String = 'NbMesDeriveEffectuees';
    S_BDR_NBMES: String = 'NbMesures';
    S_BDR_NBMESDERIVE: String = 'NbMesuresDerive';
    S_BDR_NBCYCLESDERIVE: String = 'NbCyclesDerive';
    S_BDR_DATEDEBUT: String = 'DateDebutDerive';
    S_BDR_DATENEXT: String = 'DateProchainCycleDerive';
    S_BDR_MESUREAVECFREQ: String = 'MesureAvecFrequencemetre';
    S_BDR_INIT_MANU: String = 'InitManu';

    // Chemin d'accès aux résultats
    S_EXT_EXCEL = '.xls';
    S_EXT_MODELE_XL = '.xlt';

    S_FIC_FREQ: String = 'Freq' + S_EXT_EXCEL;

    S_RACINEFIC_STAB = 'Stab';
    S_RACINEFIC_DERIV = 'Deriv';

    S_FIC_STAB: String = S_RACINEFIC_STAB + '%d' + S_EXT_EXCEL;
    S_FIC_DERIVE: String = S_RACINEFIC_DERIV + '%d' + S_EXT_EXCEL;

    S_GRAPHESTAB: String = 'GrapheStab';

    S_FIC_BESANCON: String = 'ef_utcop';
    S_PATH_SAVEBESANCON: string = 'SavBesancon';

    S_MODELE_METROLOGO = 'METROLOGO' + S_EXT_MODELE_XL;
    S_CLASSEUR_VALIDATION = 'ValidationMetrologo' + S_EXT_EXCEL;
    S_CLASSEUR_SUIVIRUBI = 'SuiviRubidiums' + S_EXT_EXCEL;

    S_NOMSHEET_DATAS_SUIVIRUBI: string = 'Donnees';
    S_NOMSHEET_GRAPH_SUIVIRUBI: string = 'Graphique';

    S_FEUILLE_VALIDATION: String = 'Validations';

    // Noms EXCEL
    S_FORMAT_DATESHEET: String = '"Le "dd/mm/yyyy';
    S_FORMAT_TIMESHEET: String = 'hh:nn:ss';
    S_FORMAT_FREQUTILISE: String = 'Fréquencemètre utilisé : %s';
    S_LIB_FNOMINALE: String = 'Fréquence nominale : ';

    S_NOMSHEET_MODELE: String = 'ModFeuille';
    S_NOMSHEET_FREQAVANT: String = 'F_Avant_Interv';
    S_NOMSHEET_FREQFINALE: String = 'F_Finale';
    //#Todo1
//    S_NOMSHEET_FREQAVANT: String = 'Fréq. avant interv.';
//    S_NOMSHEET_FREQFINALE: String = 'Fréq. finale';

    S_NOMSHEET_RECAP: String = 'Récap.';
    S_NOMSHEET_RECAP_FREQ: String = 'RécapF';
    S_NOMSHEET_RECAP_STAB: String = 'RécapS';
    S_NOMSHEET_RECAP_DERIVE: String = 'RécapD';

    S_NOM_GRAPHSTAB: String = 'GrapheStab';

    S_NOMSHEET_GRAPHE: String = 'Graph. de Dérive';

    // N° des colonnes de la feuille Frequence
    I_COLXL_HEURE: Integer = 1;
    I_COLXL_MESURE: Integer = 2;
    I_COLXL_FREELLE: Integer = 3;
    I_COLXL_DELTAF: Integer = 4;

    // N° des colonnes de la feuille recap Frequence
    I_COLXL_RECAPF_FREQNOM: Integer = 1;
    I_COLXL_RECAPF_TEMPSMES: Integer = 2;
    I_COLXL_RECAPF_FREQCORR: Integer = 3;
    I_COLXL_RECAPF_ECARTTYPE: Integer = 4;
    I_COLXL_RECAPF_FREQINDIQUEE: Integer = 5;
    I_COLXL_RECAPF_INCERTRESOL: Integer = 6;
    I_COLXL_RECAPF_INCERTSUP: Integer = 7;
    I_COLXL_RECAPF_INCERTACCRRED: Integer = 8;
    I_COLXL_RECAPF_INCERTGLOB: Integer = 9;
    I_COLXL_RECAPF_FREQFINALE: Integer = 10;

    // N° des colonnes de la feuille Recap. Stabilité
    I_COLXL_RECAPS_TEMPSGATE: Integer = 1;
    I_COLXL_RECAPS_FREQMOY: Integer = 2;
    I_COLXL_RECAPS_ECARTTYPE: Integer = 3;
    I_COLXL_RECAPS_INCERT: Integer = 4;
    I_COLXL_RECAPS_INCERTACCRED: Integer = 5;
    I_COLXL_RECAPS_INCERTGLOB: Integer = 6;
    I_COLXL_RECAPS_VALMAX: Integer = 7;
    I_COLXL_RECAPS_VALMIN: Integer = 8;
    I_COLXL_RECAPS_ETOILE: Integer = 9;

    // N° des colonnes de la feuille Recap. Dérive
    I_COLXL_RECAPD_NUMCYCLE: Integer = 1;
    I_COLXL_RECAPD_FREQMOY: Integer = 2;
    I_COLXL_RECAPD_ECARTTYPE: Integer = 3;
    I_COLXL_RECAPD_FREQLISSEE: Integer = 4;
    I_COLXL_RECAPD_INCERTGLOB: Integer = 5;
    I_COLXL_RECAPD_FMOY_FLISSEE: Integer = 6;

    // Zones nommées relatives à la feuille de validation
    ZN_VALID_CORRF_REQREF: String = 'ValidationCorrFreq';
    ZN_VALID_ECARTTYPE_INCERTGLOB: String = 'ValidationEcartTypeIncertGlob';
    ZN_VALID_MOYENNE: String = 'ValidationValMoyenne';
    ZN_VALID_MULT_FREQ: String = 'ValidationMultiplicateurFreq';
    ZN_VALID_FREQ_LOW: String = 'ValidationFreqBasses';
    ZN_VALID_FREQ_HIGH: String = 'ValidationFreqHautes';
    ZN_VALID_INCERT_RESOL: String = 'ValidationIncertResol';
    ZN_VALID_INCERT_SUPP: String = 'ValidationIncertSupp';
    ZN_VALID_INCERT_RESOL_SUPP: String = 'ValidationIncertResolEtSupp';

    AT_ZNXL: Array[enZonesNommees] Of T_ZNXL =
    (
        (Nom: 'ZNNoFiche'; bUnique: False),
        (Nom: 'ZNNoFicheRecapFreq'; bUnique: True),
        (Nom: 'ZNNoFicheRecapStab'; bUnique: True),
        (Nom: 'ZNNoFicheRecapDerive'; bUnique: True),
        (Nom: 'ZNDate'; bUnique: False),
        (Nom: 'ZNDateRecapFreq'; bUnique: True),
        (Nom: 'ZNDateRecapStab'; bUnique: True),
        (Nom: 'ZNDateRecapDerive'; bUnique: True),
        (Nom: 'ZNTypeMesure'; bUnique: False),
        (Nom: 'ZNRubidium'; bUnique: False),
        (Nom: 'ZNGate'; bUnique: False),
        (Nom: 'ZNLibGate'; bUnique: False),
        (Nom: 'ZNModeMesure'; bUnique: False),
        (Nom: 'ZNCoeffMult'; bUnique: False),
        (Nom: 'ZNEcartType'; bUnique: False),
        (Nom: 'ZNIncertEcartType'; bUnique: False),
        (Nom: 'ZNFreqRef'; bUnique: False),
        (Nom: 'ZNFreqCorr'; bUnique: False),
        (Nom: 'ZNIncertResol'; bUnique: False),
        (Nom: 'ZNIncertSup'; bUnique: False),
        (Nom: 'ZNIncertAccreditee'; bUnique: False),
        (Nom: 'ZNIncertGlobale'; bUnique: False),
        (Nom: 'ZNFreqUtilise'; bUnique: False),
        (Nom: 'ZNFNominale'; bUnique: False),
        (Nom: 'ZNValFNominale'; bUnique: False),
        (Nom: 'ZNFreqMoyReel'; bUnique: False),
        (Nom: 'ZNNbMesures'; bUnique: False),
        (Nom: 'ZNRecapF_NbMes'; bUnique: True),
        (Nom: 'ZNVariance'; bUnique: False),
        (Nom: 'ZNRecapD_NbMesCycle'; bUnique: True),
        (Nom: 'ZNRecapD_TpsMes'; bUnique: True),
        (Nom: 'ZNRecapD_TpsCycle'; bUnique: True),
        (Nom: 'ZNRecapF_DebZone'; bUnique: True),
        (Nom: 'ZNRecapF_Ligne0'; bUnique: True),
        (Nom: 'ZNRecapS_DebZone'; bUnique: True),
        (Nom: 'ZNRecapS_Ligne0'; bUnique: True),
        (Nom: 'ZNRecapD_DebZone'; bUnique: True),
        (Nom: 'ZNRecapD_Ligne0'; bUnique: True),
        (Nom: 'ZNCoeffA'; bUnique: False),
        (Nom: 'ZNCoeffB'; bUnique: False),
        (Nom: 'ZNValGateSecondes'; bUnique: False),
        (Nom: 'ZNNbMesAccredite'; bUnique: False),
        (Nom: 'ZNTempsMesureAccredite'; bUnique: False),
        (Nom: 'ZNCoeffNbMes'; bUnique: False),
        (Nom: 'ZNCoeffTempsMesure'; bUnique: False),
        (Nom: 'ZNIncertGlobNONStab'; bUnique: False),
        (Nom: 'ZNGrandeValeur'; bUnique: True),
        (Nom: 'ZNPetiteValeur'; bUnique: True),
        (Nom: 'ZNCoeffDeriveJournaliere'; bUnique: True)
        );

    AT_ZONESVALIDATION: Array[enZNValidation] Of T_ZONESVALIDATION =
    (
        (ZNLibelle: 'LibCorrfreq'; ZNDonnees: 'ValidationCorrFreq'),
        (ZNLibelle: 'LibFreqMoy'; ZNDonnees: 'ValidationValMoyenne'),
        (ZNLibelle: 'LibETIncertGlob'; ZNDonnees: 'ValidationETIncertGlob'),
        (ZNLibelle: 'LibMultFreq'; ZNDonnees: 'ValidationMultiplicateurFreq'),
        (ZNLibelle: 'LibFreqLow'; ZNDonnees: 'ValidationFreqBasses'),
        (ZNLibelle: 'LibFreqHigh'; ZNDonnees: 'ValidationFreqHautes'),
        (ZNLibelle: 'LibIncertResol'; ZNDonnees: 'ValidationIncertResol'),
        (ZNLibelle: 'LibIncertSupp'; ZNDonnees: 'ValidationIncertSupp'),
        (ZNLibelle: 'LibIncertResolEtSupp'; ZNDonnees: 'ValidationIncertResolEtSupp')
        );

    S_ZNMESURE1: String = 'ZNMesure1';
    S_ZNMESURE2: String = 'ZNMesure2';
    S_ZNPOINTINSERT: String = 'ZNPointInsertion';
    S_ZNMOYENNE: String = 'ZNPlageFReelle';
    S_ZNDELTAF: String = 'ZNDeltaF';

    S_ZNLIGNE_MODELE: String = 'ZNLigneModele';
    S_ZNCELL_DEB_COLLAGE: String = 'ZNCellDebCollage';
    S_ZNPREMIERE_CELLULE: string = 'ZNPremiereCellule';
    S_ZNTABLEAU: string = 'ZNTableau';
    S_ZNRUBI_NOM: string = 'ZNNomRubidium';
    S_ZNRUBI_DATE: string = 'ZNDateGrapheSuiviRubi';
    S_ZNRUBI_NUMSEM: string = 'ZNNumSemaine';
    S_ZNRUBI_NBSEM: string = 'ZNNbSemaines';
    

    S_MACRO_WB = 'Metrologo.xla';
//    S_MACRO_VARIANCE = S_MACRO_WB + '!Cal_Variance';
    S_MACRO_VARIANCE = 'Cal_Variance';
    S_MACRO_ECARTTYPE = 'Cal_ecart_type';
    S_MACRO_INCERTECARTTYPE = 'Cal_incert_ecart_type';
    S_MACRO_CORRFREQ = 'Cal_freq_corrigee';

    S_FORMULE_MOYENNE: String = '=IF(ISBLANK(ZNNbMesures),,AVERAGE(%s))';
    S_FORMULE_VARIANCE: String = '=Cal_variance(SUMSQ(%s),ZNNbMesures,ZNFreqMoyReel)';
    S_FORMULE_GRANDEVALEUR: String = '=LARGE(R%dC:R%dC,1)';
    S_FORMULE_PETITEVALEUR: String = '=SMALL(R%dC:R%dC,1)';

    // Chaines diverses
    S_MDPVALIDATION: String = 'METROL';
    S_MDPINCERT: String = '1135';

    AS_CDESMUX: Array[0..6] Of String = ('A1', 'A2', 'A3', 'A4' + #10 + 'B1', 'A4' + #10 + 'B2', 'A4' + #10 + 'B3', 'A4' + #10 + 'B4');

    S_LIBEL_FREQ = 'Fréquence';
    S_LIBEL_STAB = 'Stabilité';
    S_LIBEL_DERIVE = 'Dérive';
    S_LIBEL_INTERVAL = 'Intervalle';
    S_LIBEL_TACHYCONTACTS = 'Tachy. contacts';
    S_LIBEL_TACHYOPTIQUE = 'Tachy. optique';
    S_LIBEL_STROBOSCOPE = 'Stroboscope';

    S_LIBEL_TPSCOMPTAGE: String = 'Temps de comptage =';

    AS_LIBEL_TYPEMESURE: Array[enTypesMesures] Of String = (S_LIBEL_FREQ, S_LIBEL_STAB, S_LIBEL_DERIVE, S_LIBEL_FREQ, S_LIBEL_FREQ, S_LIBEL_INTERVAL,
        S_LIBEL_TACHYCONTACTS, S_LIBEL_TACHYOPTIQUE, S_LIBEL_STROBOSCOPE);

    AS_LIBEL_MODESMESURES: Array[enModeMesures] Of String = ('Mode direct', 'Coeff. multiplicateur d''écart =');

    // Requetes
    S_REQ_ASERI_FI: string = 'Select top 1 AffID from SIA..tAffaire where AffNoFI=%s';
    S_REQ_RUBACTIF: String = 'Select * from TR_METROLOGO_RUBIDIUMS Where RUB_ACTIF=1';

    S_REQ_RUBIDIUMS: String = 'Select RUB_ID, RUB_ACTIF, RUB_DESIGNATION from TR_METROLOGO_RUBIDIUMS';

    S_REQINCERT_FREQ: String = 'select A.APP_ID, A.APP_NOM as Nom, P.PLG_ID, P.PLG_LIBELLE as Plage, INCF_COEFFA as CoeffA, INCF_COEFFB as CoeffB '
    + 'from T_METROLOGO_INCERT_FREQ F '
        + 'inner join TR_METROLOGO_APPS A on A.APP_ID=F.APP_ID '
        + 'inner join TR_METROLOGO_PLAGES_FREQAPP P on P.PLG_ID=F.PLG_ID '
        + 'where F.RUB_ID = %d and F.RUB_AVECGPS = %d and F.APP_ID = %d and F.PLG_ID=%d '
        + 'order by A.APP_ID, P.PLG_ID';

    S_REQINCERT_STAB: String = 'select A.APP_ID, A.APP_NOM, G.GATE_ID, G.GATE_LIBELLE, INCS_INCERT as CoeffA '
    + 'from T_METROLOGO_INCERT_STAB S '
        + 'inner join TR_METROLOGO_APPS A on A.APP_ID=S.APP_ID '
        + 'inner join TR_METROLOGO_GATES G on G.GATE_ID=S.GATE_ID '
        + 'where S.APP_ID=%d and S.GATE_ID=%d '
        + 'order by A.APP_ID, G.GATE_ID';

    S_REQ_INCERTAUTRES: String = 'Select AUT_COEFFA as CoeffA, AUT_COEFFB as CoeffB from T_METROLOGO_INCERT_AUTRESMESURES where TYP_ID = %d';

    S_REQ_DISPINCERT_FREQ: String = 'select P.PLG_LIBELLE as [ ], case when F.RUB_AVECGPS = 0 then ''Allouis'' else ''GPS'' end as [Raccord.], '
    + 'INCF_COEFFA as a, INCF_COEFFB as b '
        + 'from T_METROLOGO_INCERT_FREQ F '
        + 'inner join TR_METROLOGO_APPS A on A.APP_ID=F.APP_ID '
        + 'inner join TR_METROLOGO_PLAGES_FREQAPP P on P.PLG_ID=F.PLG_ID '
        + 'where F.RUB_ID = %d and F.APP_ID = %d '
        + 'order by F.RUB_AVECGPS, P.PLG_ID';

    S_REQ_DISPINCERT_STAB: String = 'select G.GATE_LIBELLE as [Temps de mesure], INCS_INCERT as [Valeur] '
    + 'from T_METROLOGO_INCERT_STAB S '
        + 'inner join TR_METROLOGO_APPS A on A.APP_ID=S.APP_ID '
        + 'inner join TR_METROLOGO_GATES G on G.GATE_ID=S.GATE_ID '
        + 'where S.APP_ID=%d '
        + 'order by G.GATE_ID';

    S_REQ_DISPINCERT_AUTRES: String = 'Select AUT_COEFFA as CoeffA, AUT_COEFFB as CoeffB, AUT_LIBELLE as Libelle from T_METROLOGO_INCERT_AUTRESMESURES where TYP_ID = %d';

    S_REQINFOS_INCERT: String = 'select * from dbo.T_METROLOGO_DATAS_INCERTITUDE';

    S_REQ_DATES_GRAPHES: String = 'select * from dbo.TJ_METROLOGO_SUIVIRUBI '
	        + 'where RUB_ID=%d and RUB_AVECGPS=%d and DAT_ID between %d and %d and SUV_ECARTF is not null '
	        + 'order by DAT_ID';

    // Valeurs de frequence
    F_FREQNOMINALE: double = 10000000.0;

Var
    ListeFrequencemetres: TList;                                                // Liste des objets 'Fréquencemètre'
    ListeDerivesEnCours: TList;                                                 // Liste des dérives en cours

    dFrequenceDeReference: Double;                                              // Valeur de la fréquence du rubidium actif

    ModeMetrologo: emModesMetrologo;                                            // Flag indiquant dans quel mode fonctionne Metrologo
    OLEFalse, OLETrue: OleVariant;

    sPathResult: string;
    sPathModeles: string;
    sPathDonFI: string;

    sModeleMetrologo: string;
    sClasseurValidation: string;
    sClasseurSuiviRubi: string;

Implementation

Const
    S_SEP_POINTDEC: String = '.';
    S_SEP_VIRGDEC: String = '.';

{ TAppareilIEEE }

{*-------------------------------------------------------------------------------
  Procedure: TAppareilIEEE.Lecture
  @Author:   cb le 13/08/2009
  @Param:    spCde: String
  @Result:   string
-------------------------------------------------------------------------------}
Function TAppareilIEEE.Lecture(opMes: TMesure): Double;
Var
    blTimeOut: Boolean;
    slVal: String;

Begin

{$IFDEF Debug}
    Result := 10000000.0 + (Random / 1000.0);
    Exit;
{$ENDIF}

    if (Self.sNom = S_NOMS_APPIEE[eaRacal]) and (opMes.TypeMesure = etInterval) then begin
        slVal := MesureIntervalleRacalDana(Self.iAdresse, Self.iWriteTerm, Self.iReadTerm);
        blTimeOut := False;
    end
    else
        slVal := EcritureLectureIEEE(Self.sExeMesure, Self.iAdresse, Self.iWriteTerm, Self.iReadTerm, Self.bGereSRQ, blTimeOut);
//    slVal := EcritureLectureIEEE(spCde, Self.iAdresse, Self.iWriteTerm, Self.iReadTerm, Self.bGereSRQ, blTimeOut);


    If blTimeOut Then
        Result := 0.0
    Else
        Result := StrToFloat(StringReplace(Mid(slVal , Self.iTailleHeaderReponse), S_SEP_POINTDEC, S_SEP_VIRGDEC, []));

End;

{ TMesure }

{*-------------------------------------------------------------------------------
  Procedure: TMesure.setNumFI
  @Author:   cb le 28/08/2009
  @Param:    Const Value: String
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TMesure.setNumFI(Const Value: String);
Begin

    If Trim(Value) = '' Then exit;

    FNumFI := Value;
    FNumFISlash := StringReplace(FNumFI, '_', '/', []);
    FID_FI := NumDocE2MToID(FNumFISlash);

    FDossierResultats := Format(sPathResult, [FNumFI]);
    FNomFicFreq := S_FIC_FREQ;                                                  // Définit nom du fichier résultat Freq
    FNomFicStab := GetIndexedFile(S_RACINEFIC_STAB);                            // Définit nom du fichier résultat Stab
    FNomFicDerive := GetIndexedFile(S_RACINEFIC_DERIV);                         // Définit nom du fichier résultat Stab

End;

{*-------------------------------------------------------------------------------
  Procedure: TMesure.Clear
  @Author:   cb le 01/09/2009
  @Param:    None
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TMesure.Init();
Begin

    With Self Do Begin
        FNominale := F_FREQNOMINALE;
        bInitManu := False;
        VoieMux := I_MUX_V1;
        sNumFI := '';
        Frequencemetre := eaStanford;                                           // Freq. par défaut = Stanford 620
        ModeMesure := emDirect;                                                 // Mode direct
        TypeMesure := etFrequence;
        IndexMultiplicateur := I_MULT_COEF4;                                    // Valeur par défaut du multiplicateur de fréquence
        InputStanford := eiStanfordLowZ;
        CouplingStanford := eiAC;
        InputRacal := eiInputA50;
        CouplingRacal := eiAC;
        InputEIP := eiBand3;
        IndexIntervalDerive := I_IDX_PERIODICITE_DERIVE_DEFAUT;
        NbMesures := I_NBMESURES_DEFAUT;
        NbMesDerive := I_NBMESDERIVE_DEFAULT;
        NbCyclesDerive := I_NBCYCLESDERIVE_DEFAUT;
        DateDebutDerive := Now;
        DateNextCycleDerive := DateDebutDerive;
        NbCyclesDeriveEffectuees := 0;
        bMesureAvecFreq := False;

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TMesure.IsFreqMeasurement
  @Author:   cb le 07/09/2009
  @Param:    None
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TMesure.IsFreqMeasurement: Boolean;
Begin

    Result := TypeMesure In [etFreqAvantInterv, etFrequence, etFreqFinale, etInterval, etTachyContact, etTachyOptique, etStroboscope];

End;

{*-------------------------------------------------------------------------------
-------------------------------------------------------------------------------}

{*-------------------------------------------------------------------------------
  Procedure: TMesure.GetIndexFile
  @Author:   cb le 07/09/2009
  @Param:    spRacineNom: String
  @Result:   String
-------------------------------------------------------------------------------}
Function TMesure.GetIndexedFile(spRacineNom: String): String;
Var
    slName: String;
    ilIndex: Integer;
    blOk: Boolean;

Begin

    Result := '';
    blOk := False;

    With TJvSearchFiles.Create(Nil) Do
    Begin
        DirOption := doExcludeSubDirs;
        FileParams.FileMask := spRacineNom + '*' + S_EXT_EXCEL;
        FileParams.SearchTypes := [stFileMask];
        Options := [soSearchFiles, soSorted];
        RootDirectory := FDossierResultats;

        If Search And (TotalFiles > 0) Then
        Begin
            slName := ChangeFileExt(ExtractFileName(Files[ToTalFiles - 1]), '');
            If Length(slName) > Length(spRacineNom) Then
            Begin
                slName := Mid(slName, Length(spRacineNom) + 1);
                If TryStrToInt(slName, ilIndex) Then
                Begin
                    If Self.TypeMesure = etDerive Then
                    Begin                                                       // Si Dérive, ne génère un nouveau nom que si le nombre de cycles de dérive effectué = 0
                        If Self.NbCyclesDeriveEffectuees = 0 Then
                            Inc(ilIndex);
                        Result := spRacineNom + IntToStr(ilIndex) + S_EXT_EXCEL
                    End
                    Else
                        Result := spRacineNom + IntToStr(ilIndex + 1) + S_EXT_EXCEL;

                    blOk := True;
                End;
            End;
        End;
        If Not blOk Then
            Result := spRacineNom + '1' + S_EXT_EXCEL;

        Free;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TMesure.Create
  @Author:   cb le 30/09/2009
  @Param:    None
  @Result:   None
-------------------------------------------------------------------------------}
Constructor TMesure.Create;
Begin

    Inherited;
    TimerDerive := TTimer.Create(Application);
    TimerDerive.OnTimer := TimerEllapsed;
    TimerDerive.Enabled := False;

End;

{*-------------------------------------------------------------------------------
  Procedure: TMesure.TimerEllapsed
  @Author:   cb le 30/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TMesure.TimerEllapsed(Sender: TObject);
Begin

    If Assigned(FTimeDerive) Then
        FTimeDerive(Self);

End;

{*-------------------------------------------------------------------------------
  Procedure: TMesure.ReadFromRegistry
  @Author:   cb le 30/09/2009
  @Param:    spNoFI: String
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TMesure.ReadFromRegistry(): Boolean;
Var
    slClef: String;
    ilTempInt: Integer;

Begin

    With TRegistry.Create Do
    Begin
        RootKey := HKEY_LOCAL_MACHINE;
        slClef := Format(S_KEY_BDR, [Self.sNumFI]);
        Result := False;

        If OpenKey(slClef, True) Then Begin                                     // Si OpenKey OK
            Result := True;
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_FNOMINALE, Self.FNominale));
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_NBMESACCRED, Self.iNbMesAccred));
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_TEMPSACCRED, Self.iTmpsMesAccred));
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_VOIEMUX, Self.VoieMux));
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_FREQ, ilTempInt));
            Self.Frequencemetre := enAppareilsIEEE(ilTempInt);
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_MODEMESURE, ilTempInt));
            Self.ModeMesure := enModeMesures(ilTempInt);
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_IDXMULT, Self.IndexMultiplicateur));
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_TYPEMES, ilTempInt));
            Self.TypeMesure := enTypesMesures(ilTempInt);
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_INPUTSR620, ilTempInt));
            Self.InputStanford := enInputStanford(ilTempInt);
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_INPUT1996, ilTempInt));
            Self.InputRacal := enInputRacal(ilTempInt);
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_INPUT545, ilTempInt));
            Self.InputEIP := enInputEIP(ilTempInt);
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_CPLGSR620, ilTempInt));
            Self.CouplingStanford := enInputCoupling(ilTempInt);
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_CPLG1996, ilTempInt));
            Self.CouplingRacal := enInputCoupling(ilTempInt);
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_INDEXGATE, Self.IndexGate));
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_IDX_INTERVDERIVE, Self.IndexIntervalDerive));
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_NBMES_DERIVES_EFFECTUEES, Self.NbCyclesDeriveEffectuees));
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_NBMES, Self.NbMesures));
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_NBMESDERIVE, Self.NbMesDerive));
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_NBCYCLESDERIVE, Self.NbCyclesDerive));
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_DATEDEBUT, Double(Self.DateDebutDerive)));
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_DATENEXT, Double(Self.DateNextCycleDerive)));
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_MESUREAVECFREQ, Self.bMesureAvecFreq));
            Result := Result And (ReadRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_INIT_MANU, Self.bInitManu));

            If Not Result Then
                DeleteKey(slClef);

            CloseKey;
            Free;
        End;

        If Not Result Then
            Application.MessageBox(PChar('Erreur lors de la récupération des informations de dérive dans la Base de registre !'),
                'Attention', MB_OK + MB_ICONERROR + MB_TOPMOST);

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TMesure.SaveToRegistry
  @Author:   cb le 30/09/2009
  @Param:    None
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TMesure.SaveToRegistry;
Var
    slClef: String;
    blOk: Boolean;

Begin

    With TRegistry.Create Do
    Begin
        RootKey := HKEY_LOCAL_MACHINE;
        slClef := Format(S_KEY_BDR, [Self.sNumFI]);
        blOk := False;

        If OpenKey(slClef, True) Then Begin                                     // Si OpenKey OK
            blOk := True;
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_FNOMINALE, Self.FNominale));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_NBMESACCRED, Self.iNbMesAccred));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_TEMPSACCRED, Self.iTmpsMesAccred));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_VOIEMUX, Self.VoieMux));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_FREQ, Ord(Self.Frequencemetre)));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_MODEMESURE, Ord(Self.ModeMesure)));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_IDXMULT, Self.IndexMultiplicateur));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_TYPEMES, Ord(Self.TypeMesure)));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_INPUTSR620, Ord(Self.InputStanford)));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_INPUT1996, Ord(Self.InputRacal)));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_INPUT545, Ord(Self.InputEIP)));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_CPLGSR620, Ord(Self.CouplingStanford)));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_CPLG1996, Ord(Self.CouplingRacal)));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_INDEXGATE, Self.IndexGate));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_IDX_INTERVDERIVE, Self.IndexIntervalDerive));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_NBMES_DERIVES_EFFECTUEES, Self.NbCyclesDeriveEffectuees));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_NBMES, Self.NbMesures));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_NBMESDERIVE, Self.NbMesDerive));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_NBCYCLESDERIVE, Self.NbCyclesDerive));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_DATEDEBUT, TDateTime(Self.DateDebutDerive)));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_DATENEXT, Self.DateNextCycleDerive));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_MESUREAVECFREQ, Self.bMesureAvecFreq));
            blOk := blOk And (WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_INIT_MANU, Self.bInitManu));

            If Not blOk Then
                DeleteKey(slClef);

            CloseKey;
            Free;
        End;

        If Not blOk Then
            Application.MessageBox(PChar('Erreur lors de l''inscription des informations de dérive dans la Base de registre !' + CRLF
                + 'Si Metrologo n''est pas fermé, ceci est sans incidence; dans le cas contraire cette mesure de dérive ne sera pas relancée au redémarrage de Metrologo !'),
                'Attention', MB_OK + MB_ICONWARNING + MB_TOPMOST);

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TMesure.UpdateRegistry
  @Author:   cb le 30/09/2009
  @Param:    None
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TMesure.UpdateRegistry;
Var
    slClef: String;

Begin

    With TRegistry.Create Do
    Begin
        RootKey := HKEY_LOCAL_MACHINE;
        slClef := Format(S_KEY_BDR, [Self.sNumFI]);

        If OpenKey(slClef, True) Then Begin                                     // Si OpenKey OK
            WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_NBMES_DERIVES_EFFECTUEES, Self.NbCyclesDeriveEffectuees);
            WriteRegistryValue(HKEY_LOCAL_MACHINE, slClef, S_BDR_DATENEXT, Self.DateNextCycleDerive);

            CloseKey;
            Free;
        End;
    End;

End;

{*------------------------------------------------------------------------------
  Procedure: TMesure.RemoveFromRegistry
  @Author:   cb le 01/10/2009
  @Param:    None
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TMesure.RemoveFromRegistry;
Var
    slClef: String;

Begin

    With TRegistry.Create Do Begin
        RootKey := HKEY_LOCAL_MACHINE;
        slClef := Format(S_KEY_BDR, [Self.sNumFI]);
        DeleteKey(slClef);
        CloseKey;
    End;

End;


{*------------------------------------------------------------------------------
  Procedure: TAppareilIEEE.MesureIntervalleRacalDana
  @Author:   cb le 15/03/2011
  @Param:    ipAdresse, ipWriteTerm, ipReadTerm
  @Result:   string
-------------------------------------------------------------------------------}
function TAppareilIEEE.MesureIntervalleRacalDana(ipAdresse: Smallint; ipWriteTerm, ipReadTerm: Integer): string;
var
    blLoop, blTO: Boolean;
    ilStat: Smallint;

begin

    EcritureIEEE('QM0', ipAdresse, ipWriteTerm);            // S'assure que l'on ne génère pas de SRQ

    blLoop := True;
    While blLoop do begin
        ReadStatusByte(GPIB0, ipAdresse, ilStat);
        if ilStat and 16 = 16 then
            LectureIEEE(ipAdresse, ipReadTerm, blTO)
        else
            blLoop := False;
    end;

    EcritureIEEE('RE', ipAdresse, ipWriteTerm);             // S'assure que l'on ne génère pas de SRQ
    SetRemoteLocal(ipAdresse, False);

    blLoop := True;
    While blLoop do begin
        ReadStatusByte(GPIB0, ipAdresse, ilStat);
        blLoop := (ilStat and 16) <> 16;
        if blLoop then begin
            Sleep(1000);
            Application.ProcessMessages;
        end;
    end;

    Result := LectureIEEE(ipAdresse, ipReadTerm, blTO);

end;


//----------------------------- Initialization -----------------------------------

Initialization
    OLEFalse := False;
    OLETrue := True;

End.
