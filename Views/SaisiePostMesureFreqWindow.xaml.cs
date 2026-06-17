using Metrologo.ViewModels;
using System;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class SaisiePostMesureFreqWindow : FluentWindow
    {
        // On veut repasser au premier plan même quand Excel (un autre process) garde le focus.
        // En pratique Topmost ne suffit pas toujours : juste après qu'Excel a pris la main via
        // Interop, Windows verrouille le foreground. La parade classique : attacher notre thread
        // d'input au sien le temps d'appeler SetForegroundWindow, ce qui lève ce verrou.
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
