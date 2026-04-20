using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class ModifValBesanconWindow : FluentWindow
    {
        public ModifValBesanconViewModel ViewModel { get; }

        public ModifValBesanconWindow() : this(new ModifValBesanconViewModel()) { }

        public ModifValBesanconWindow(ModifValBesanconViewModel vm)
        {
            InitializeComponent();
            ViewModel = vm;
            DataContext = ViewModel;
            ViewModel.CloseAction = result => { DialogResult = result; Close(); };
        }
    }
}
