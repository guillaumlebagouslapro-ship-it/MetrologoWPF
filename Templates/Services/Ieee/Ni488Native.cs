using System.Runtime.InteropServices;

namespace Metrologo.Services.Ieee
{
    /// <summary>
    /// Wrapper P/Invoke des fonctions NI-488.2 natives (ni4882.dll). Approche
    /// utilisée historiquement par le Delphi (via dpib32) et qui contourne la couche
    /// NI-VISA managée — gain typique : write+read GPIB ~190 ms → ~30-80 ms.
    /// </summary>
    internal static class Ni488Native
    {
        private const string Dll = "ni4882.dll";

        // ibdev : ouvre un handle (ud) vers un appareil
        //   boardID  : index de la carte GPIB (0 pour GPIB0)
        //   pad      : adresse primaire (1..30)
        //   sad      : adresse secondaire (0 = pas de SAD)
        //   tmo      : code timeout (voir constantes T*)
        //   eot      : 1 = assert EOI sur le dernier octet en écriture
        //   eos      : configuration EOS (0 = aucun caractère terminateur)
        [DllImport(Dll, EntryPoint = "ibdev", ExactSpelling = true)]
        public static extern int ibdev(int boardID, int pad, int sad, int tmo, int eot, int eos);

        // ibwrt : écrit des données vers l'appareil
        [DllImport(Dll, EntryPoint = "ibwrt", ExactSpelling = true)]
        public static extern int ibwrt(int ud, byte[] buffer, int cnt);

        // ibrd : lit des données depuis l'appareil
        [DllImport(Dll, EntryPoint = "ibrd", ExactSpelling = true)]
        public static extern int ibrd(int ud, byte[] buffer, int cnt);

        // ibclr : envoie un Selected Device Clear
        [DllImport(Dll, EntryPoint = "ibclr", ExactSpelling = true)]
        public static extern int ibclr(int ud);

        // ibrsp : serial poll, retourne l'octet de statut
        [DllImport(Dll, EntryPoint = "ibrsp", ExactSpelling = true)]
        public static extern int ibrsp(int ud, out byte spr);

        // ibonl : ud online (v=1) ou offline (v=0, ferme le handle)
        [DllImport(Dll, EntryPoint = "ibonl", ExactSpelling = true)]
        public static extern int ibonl(int ud, int v);

        // ibconfig : configure une option du handle. Sert pour timeout, EOS, etc.
        // Sur les versions récentes de NI-488.2 (ni4882.dll), ibtmo et ibeos ne sont
        // PAS exportés directement — il faut passer par ibconfig avec la bonne constante.
        [DllImport(Dll, EntryPoint = "ibconfig", ExactSpelling = true)]
        public static extern int ibconfig(int ud, int option, int value);

        // Constantes ibconfig (extraites de ni4882.h)
        public const int IbcTMO     = 0x0003;  // Timeout setting (équivalent ibtmo)
        public const int IbcEOSrd   = 0x000C;  // End on EOS char during read
        public const int IbcEOSchar = 0x000F;  // EOS character
        public const int IbcEOSbits = 0x000B;  // EOS mode bits (TERM, REOS, BIN, XEOS)

        public static int ibtmo(int ud, int tmo) => ibconfig(ud, IbcTMO, tmo);
        public static int ibeos(int ud, int v)
        {
            // v = REOS | char (notre convention) → on configure char + bits séparément
            int car = v & 0xFF;
            int bits = (v >> 8);  // REOS = 0x0400 → 4 → bit 0x04 (REOS_BIT) — voir ni4882.h
            ibconfig(ud, IbcEOSchar, car);
            return ibconfig(ud, IbcEOSbits, bits);
        }

        // SendIFC : envoie un Interface Clear sur le bus (broadcast)
        [DllImport(Dll, EntryPoint = "SendIFC", ExactSpelling = true)]
        public static extern void SendIFC(int boardID);

        // Variables thread-locales (sta = status, err = code erreur, cntl = octets transférés)
        // Lues après chaque appel ib* pour diagnostic.
        [DllImport(Dll, EntryPoint = "ThreadIbsta", ExactSpelling = true)]
        public static extern int ThreadIbsta();

        [DllImport(Dll, EntryPoint = "ThreadIberr", ExactSpelling = true)]
        public static extern int ThreadIberr();

        // Note : pas de ThreadIbcntl (avec 'l') — sur ni4882.dll c'est ThreadIbcnt.
        [DllImport(Dll, EntryPoint = "ThreadIbcnt", ExactSpelling = true)]
        public static extern int ThreadIbcntl();

        // Bits du registre status (ibsta)
        public const int ERR_BIT  = 0x8000;
        public const int TIMO_BIT = 0x4000;
        public const int END_BIT  = 0x2000;

        // EOS mode bits (à OR avec le caractère terminateur)
        public const int REOS = 0x0400;  // Termine le read sur le char EOS

        /// <summary>
        /// Convertit une durée en ms vers un code timeout NI-488.2 (T1ms=5 … T1000s=17).
        /// </summary>
        public static int MapTimeoutCode(int ms)
        {
            if (ms <= 0) return 0;       // TNONE
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
