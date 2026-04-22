using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class EnregistrementAppareilWindow : FluentWindow
    {
        public EnregistrementAppareilViewModel ViewModel { get; }

        public EnregistrementAppareilWindow(EnregistrementAppareilViewModel vm)
        {
            InitializeComponent();
            ViewModel = vm;
            DataContext = vm;
            vm.CloseAction = ok =>
            {
                if (IsVisible) { DialogResult = ok; Close(); }
            };
        }
    }
}
