using System.Runtime.InteropServices;

namespace Metrologo.Services.Ieee
{
    /// <summary>
    /// Couche P/Invoke vers les fonctions NI-488.2 natives (ni4882.dll). On reprend la recette du
    /// Delphi d'origine (dpib32) : en court-circuitant la couche NI-VISA managée, un cycle write+read
    /// GPIB tombe de ~190 ms à ~30-80 ms.
    /// </summary>
    internal static class Ni488Native
    {
        private const string Dll = "ni4882.dll";

        // ibdev : ouvre un handle (ud) vers un appareil
        //   boardID  : numéro de la carte GPIB (0 pour GPIB0)
        //   pad      : adresse primaire (1..30)
        //   sad      : adresse secondaire (0 si pas de SAD)
        //   tmo      : code de timeout (cf. les constantes T*)
        //   eot      : 1 = affirme EOI sur le dernier octet à l'écriture
        //   eos      : réglage EOS (0 = pas de caractère terminateur)
        [DllImport(Dll, EntryPoint = "ibdev", ExactSpelling = true)]
        public static extern int ibdev(int boardID, int pad, int sad, int tmo, int eot, int eos);

        // ibwrt : envoie des données à l'appareil
        [DllImport(Dll, EntryPoint = "ibwrt", ExactSpelling = true)]
        public static extern int ibwrt(int ud, byte[] buffer, int cnt);

        // ibrd : récupère des données depuis l'appareil
        [DllImport(Dll, EntryPoint = "ibrd", ExactSpelling = true)]
        public static extern int ibrd(int ud, byte[] buffer, int cnt);

        // ibclr : envoie un Selected Device Clear (SDC)
        [DllImport(Dll, EntryPoint = "ibclr", ExactSpelling = true)]
        public static extern int ibclr(int ud);

        // ibrsp : serial poll, renvoie l'octet de statut
        [DllImport(Dll, EntryPoint = "ibrsp", ExactSpelling = true)]
        public static extern int ibrsp(int ud, out byte spr);

        // ibonl : passe le ud online (v=1) ou offline (v=0, ce qui ferme le handle)
        [DllImport(Dll, EntryPoint = "ibonl", ExactSpelling = true)]
        public static extern int ibonl(int ud, int v);

        // ibconfig : règle une option du handle (timeout, EOS, etc.).
        // Sur les NI-488.2 récents (ni4882.dll), ibtmo et ibeos ne sont PAS exportés ;
        // on est donc obligé de passer par ibconfig avec la bonne constante.
        [DllImport(Dll, EntryPoint = "ibconfig", ExactSpelling = true)]
        public static extern int ibconfig(int ud, int option, int value);

        // Constantes ibconfig (reprises de ni4882.h)
        public const int IbcTMO     = 0x0003;  // réglage du timeout (l'équivalent de ibtmo)
        public const int IbcEOSrd   = 0x000C;  // terminer la lecture sur le caractère EOS
        public const int IbcEOSchar = 0x000F;  // le caractère EOS lui-même
        public const int IbcEOSbits = 0x000B;  // bits de mode EOS (TERM, REOS, BIN, XEOS)

        public static int ibtmo(int ud, int tmo) => ibconfig(ud, IbcTMO, tmo);
        public static int ibeos(int ud, int v)
        {
            // v = REOS | char selon notre convention ; ici on sépare le caractère des bits de mode
            int car = v & 0xFF;
            int bits = (v >> 8);  // REOS = 0x0400 -> 4 -> bit 0x04 (REOS_BIT), cf. ni4882.h
            ibconfig(ud, IbcEOSchar, car);
            return ibconfig(ud, IbcEOSbits, bits);
        }

        // SendIFC : émet un Interface Clear sur tout le bus (broadcast)
        [DllImport(Dll, EntryPoint = "SendIFC", ExactSpelling = true)]
        public static extern void SendIFC(int boardID);

        // Variables propres à chaque thread (sta = statut, err = code d'erreur, cntl = octets transférés).
        // On les relit après chaque appel ib* quand on a besoin de diagnostiquer.
        [DllImport(Dll, EntryPoint = "ThreadIbsta", ExactSpelling = true)]
        public static extern int ThreadIbsta();

        [DllImport(Dll, EntryPoint = "ThreadIberr", ExactSpelling = true)]
        public static extern int ThreadIberr();

        // attention : ni4882.dll n'a pas de ThreadIbcntl (avec un 'l') ; l'export réel s'appelle ThreadIbcnt
        [DllImport(Dll, EntryPoint = "ThreadIbcnt", ExactSpelling = true)]
        public static extern int ThreadIbcntl();

        // Bits du registre de statut (ibsta)
        public const int ERR_BIT  = 0x8000;
        public const int TIMO_BIT = 0x4000;
        public const int END_BIT  = 0x2000;

        // Bits du mode EOS (à combiner par OR avec le caractère terminateur)
        public const int REOS = 0x0400;  // arrête la lecture dès qu'on rencontre le char EOS

        /// <summary>Traduit une durée en ms vers le code de timeout NI-488.2 correspondant (T1ms=5 ... T1000s=17).</summary>
        public static int MapTimeoutCode(int ms)
        {
            if (ms <= 0) return 0;       // TNONE (pas de timeout)
            if (ms <= 1) return 5;       // T1ms
            if (ms <= 3) return 6;       // T3ms
            if (ms <= 10) return 7;      // T10ms
            if (ms <= 30) return 8;      // T30ms
            if (ms <= 100) return 9;     // T100ms
            if (ms <= 300) return 10;    // T300ms
            if (ms <= 1000) return 11;   // T1s
            if (ms <= 3000) return 12;   // T3s
            if (ms <= 10_000) return 13; // T10s
            if (ms <= 30_000) return 14; // T30s
            if (ms <= 100_000) return 15;// T100s
            if (ms <= 300_000) return 16;// T300s
            return 17;                    // T1000s
        }
    }
}
