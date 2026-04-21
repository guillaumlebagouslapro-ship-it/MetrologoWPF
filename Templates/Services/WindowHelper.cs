using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Metrologo.Services
{
    /// <summary>
    /// Helpers Win32 pour piloter une fenêtre externe (ex: Excel).
    /// </summary>
    internal static class WindowHelper
    {
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        private const uint WM_CLOSE = 0x0010;

        /// <summary>
        /// Envoie WM_CLOSE à toutes les fenêtres dont le titre contient à la fois
        /// <paramref name="nomFichier"/> et le mot « Excel ».
        /// </summary>
        /// <returns>Nombre de fenêtres notifiées.</returns>
        public static int FermerFenetresExcel(string nomFichier)
        {
            int nb = 0;
            EnumWindows((hWnd, _) =>
            {
                if (!IsWindowVisible(hWnd)) return true;

                int length = GetWindowTextLength(hWnd);
                if (length == 0) return true;

                var sb = new StringBuilder(length + 1);
                GetWindowText(hWnd, sb, sb.Capacity);
                var titre = sb.ToString();

                // Excel affiche son nom dans le titre (« Xxxx.xlsm - Excel »)
                bool contientFichier = titre.IndexOf(nomFichier, StringComparison.OrdinalIgnoreCase) >= 0;
                bool contientExcel = titre.IndexOf("Excel", StringComparison.OrdinalIgnoreCase) >= 0;

                if (contientFichier && contientExcel)
                {
                    PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                    nb++;
                }
                return true;
            }, IntPtr.Zero);
            return nb;
        }
    }
}
