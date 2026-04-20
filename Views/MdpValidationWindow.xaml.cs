using System.Windows;
using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class MdpValidationWindow : FluentWindow
    {
        public MdpValidationViewModel ViewModel { get; }

        public MdpValidationWindow() : this(new MdpValidationViewModel()) { }

        public MdpValidationWindow(MdpValidationViewModel vm)
        {
            InitializeComponent();
            ViewModel = vm;
            DataContext = ViewModel;
            ViewModel.CloseAction = result => { DialogResult = result; Close(); };
            Loaded += (_, _) => txtMdp.Focus();
        }

        private void BtnVerifier_Click(object sender, RoutedEventArgs e)
        {
            ViewModel.Verifier(txtMdp.Password);
        }
    }
}
