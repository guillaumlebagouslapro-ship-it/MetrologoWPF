using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;

namespace Metrologo.Views
{
    /// <summary>
    /// Fenêtre flottante d'arrêt d'urgence pendant une mesure. Approche maximale
    /// fiabilité :
    /// <list type="bullet">
    ///   <item>Pas de <c>AllowsTransparency</c> (cause connue de problèmes de clic sur
    ///     Topmost) — fenêtre opaque solide rouge vif, ultra-visible.</item>
    ///   <item><c>PreviewMouseLeftButtonDown</c> au niveau Window → tout clic n'importe
    ///     où sur la fenêtre déclenche l'arrêt (pas besoin de viser un bouton).</item>
    ///   <item><b>Raccourci global Ctrl+Shift+S</b> via <c>RegisterHotKey</c> — fonctionne
    ///     même si Excel monopolise la souris (passe par le système Windows, pas par
    ///     notre input).</item>
    ///   <item>Touches Échap / Espace / Entrée locales pour arrêter au clavier.</item>
    ///   <item>Hook <c>WM_SETCURSOR</c> qui force le curseur Arrow via Win32 SetCursor —
    ///     contourne l'I-beam hérité d'Excel.</item>
    /// </list>
    /// </summary>
    public partial class StopMesureFloatingWindow : Window
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        private const int  HOTKEY_ID    = 0x4D45;
        private const uint MOD_SHIFT    = 0x0004;
        private const uint MOD_CONTROL  = 0x0002;
        private const uint VK_S         = 0x53;
        private const int  WM_HOTKEY    = 0x0312;

        [DllImport("user32.dll")]
        private static extern IntPtr SetCursor(IntPtr hCursor);
        [DllImport("user32.dll")]
        private static extern IntPtr LoadCursor(IntPtr hInstance, int lpCursorName);
        private const int IDC_HAND      = 32649;
        private const int WM_SETCURSOR  = 0x0020;

        private HwndSource? _source;
        private Action? _onStop;
        private bool _stopDejaDemande;
        private bool _hotkeyEnregistre;

        public StopMesureFloatingWindow()
        {
            InitializeComponent();

            Loaded += (_, _) =>
            {
                PositionnerCoinHautDroit();
                ForcerForegroundEtFocus();
            };

            // Tout clic gauche sur la fenêtre = arrêt, peu importe où.
            // PreviewMouseLeftButtonDown = on intercepte avant que d'autres handlers
            // (drag, focus, etc.) consomment l'event.
            PreviewMouseLeftButtonDown += (_, e) =>
            {
                DeclencherStop();
                e.Handled = true;
            };

            // Échappatoire clavier locale.
            KeyDown += (_, e) =>
            {
                if (e.Key == Key.Escape || e.Key == Key.Space || e.Key == Key.Enter)
                {
                    DeclencherStop();
                    e.Handled = true;
                }
            };
        }

        public void Configurer(Action onStop) => _onStop = onStop;

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            var hWnd = new WindowInteropHelper(this).Handle;
            _source = HwndSource.FromHwnd(hWnd);
            _source?.AddHook(WndProc);

            // Ctrl+Shift+S global — fonctionne même quand Excel a le focus.
            _hotkeyEnregistre = RegisterHotKey(hWnd, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, VK_S);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            switch (msg)
            {
                case WM_HOTKEY when wParam.ToInt32() == HOTKEY_ID:
                    DeclencherStop();
                    handled = true;
                    return IntPtr.Zero;

                case WM_SETCURSOR:
                    // Force le curseur main partout sur la fenêtre — visualise clairement
                    // qu'on peut cliquer n'importe où, peu importe ce qu'Excel a mis comme curseur.
                    SetCursor(LoadCursor(IntPtr.Zero, IDC_HAND));
                    handled = true;
                    return new IntPtr(1);
            }
            return IntPtr.Zero;
        }

        protected override void OnClosed(EventArgs e)
        {
            if (_hotkeyEnregistre)
            {
                var hWnd = new WindowInteropHelper(this).Handle;
                if (hWnd != IntPtr.Zero) UnregisterHotKey(hWnd, HOTKEY_ID);
            }
            _source?.RemoveHook(WndProc);
            base.OnClosed(e);
        }

        private void DeclencherStop()
        {
            if (_stopDejaDemande) return;
            _stopDejaDemande = true;
            TxtBouton.Text = "⏳ ARRÊT EN COURS...";
            Mouse.OverrideCursor = null;
            try { _onStop?.Invoke(); } catch { /* swallow */ }
        }

        private void PositionnerCoinHautDroit()
        {
            var screen = System.Windows.SystemParameters.WorkArea;
            Left = screen.Right - Width - 20;
            Top  = screen.Top + 20;
        }

        private void ForcerForegroundEtFocus()
        {
            var hWnd = new WindowInteropHelper(this).Handle;
            if (hWnd != IntPtr.Zero)
            {
                var fgHwnd = GetForegroundWindow();
                uint fgThread = GetWindowThreadProcessId(fgHwnd, out _);
                uint curThread = GetCurrentThreadId();
                if (fgThread != curThread)
                {
                    AttachThreadInput(fgThread, curThread, true);
                    SetForegroundWindow(hWnd);
                    AttachThreadInput(fgThread, curThread, false);
                }
                else
                {
                    SetForegroundWindow(hWnd);
                }
            }
            Activate();
            Focus();
        }
    }
}
