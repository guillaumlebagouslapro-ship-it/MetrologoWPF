using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class SaisieValFreqWindow : FluentWindow
    {
        public SaisieValFreqViewModel ViewModel { get; }

        public SaisieValFreqWindow() : this(new SaisieValFreqViewModel()) { }

        public SaisieValFreqWindow(SaisieValFreqViewModel vm)
        {
            InitializeComponent();
            ViewModel = vm;
            DataContext = ViewModel;
            ViewModel.CloseAction = result => { DialogResult = result; Close(); };
            Loaded += (_, _) => txtValeur.Focus();
        }
    }
}
