using Metrologo.ViewModels;
using System;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class SaisiePostMesureFreqWindow : FluentWindow
    {
        // Force le foreground même si Excel (autre process) tient le focus.
        // Topmost ne suffit pas toujours quand Excel vient juste de prendre le
        // foreground via Interop : on attache le thread input et on appelle
        // SetForegroundWindow pour passer outre la protection foreground-lock.
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

        public SaisiePostMesureFreqViewModel ViewModel { get; }

        public SaisiePostMesureFreqWindow() : this(new SaisiePostMesureFreqViewModel()) { }

        public SaisiePostMesureFreqWindow(SaisiePostMesureFreqViewModel vm)
        {
            InitializeComponent();
            ViewModel = vm;
            DataContext = ViewModel;
            ViewModel.CloseAction = result => { DialogResult = result; Close(); };
            Loaded += (_, _) =>
            {
                ForcerForeground();
                Activate();
                txtFreq.Focus();
            };
        }

        private void ForcerForeground()
        {
            var hWnd = new WindowInteropHelper(this).Handle;
            if (hWnd == IntPtr.Zero) return;

            var foregroundHwnd = GetForegroundWindow();
            uint foregroundThread = GetWindowThreadProcessId(foregroundHwnd, out _);
            uint currentThread = GetCurrentThreadId();

            if (foregroundThread != currentThread)
            {
                AttachThreadInput(foregroundThread, currentThread, true);
                SetForegroundWindow(hWnd);
                AttachThreadInput(foregroundThread, currentThread, false);
            }
            else
            {
                SetForegroundWindow(hWnd);
            }
        }
    }
}
