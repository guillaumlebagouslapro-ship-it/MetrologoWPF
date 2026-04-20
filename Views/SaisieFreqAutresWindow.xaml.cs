using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class SaisieFreqAutresWindow : FluentWindow
    {
        public SaisieFreqAutresViewModel ViewModel { get; }

        public SaisieFreqAutresWindow() : this(new SaisieFreqAutresViewModel()) { }

        public SaisieFreqAutresWindow(SaisieFreqAutresViewModel vm)
        {
            InitializeComponent();
            ViewModel = vm;
            DataContext = ViewModel;
            ViewModel.CloseAction = result => { DialogResult = result; Close(); };
        }
    }
}
