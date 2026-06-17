using Metrologo.ViewModels;
using System;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class SaisiePostMesureTachyWindow : FluentWindow
    {
        // On reprend les mêmes P/Invoke que SaisiePostMesureFreqWindow pour réussir à passer
        // devant Excel une fois qu'il s'est mis au premier plan via Interop.
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

        public SaisiePostMesureTachyViewModel ViewModel { get; }

        public SaisiePostMesureTachyWindow() : this(new SaisiePostMesureTachyViewModel()) { }

        public SaisiePostMesureTachyWindow(SaisiePostMesureTachyViewModel vm)
        {
            InitializeComponent();
            ViewModel = vm;
            DataContext = ViewModel;
            ViewModel.CloseAction = result => { DialogResult = result; Close(); };
            Loaded += (_, _) =>
            {
                ForcerForeground();
                Activate();
                txtResolution.Focus();
                txtResolution.SelectAll();
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
