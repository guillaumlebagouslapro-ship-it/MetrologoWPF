using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class SaisieMarcheHebdoWindow : FluentWindow
    {
        public SaisieMarcheHebdoViewModel ViewModel { get; }

        public SaisieMarcheHebdoWindow() : this(new SaisieMarcheHebdoViewModel()) { }

        public SaisieMarcheHebdoWindow(SaisieMarcheHebdoViewModel vm)
        {
            InitializeComponent();
            ViewModel = vm;
            DataContext = ViewModel;
            ViewModel.CloseAction = result => { DialogResult = result; Close(); };
        }
    }
}
