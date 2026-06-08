using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class BesanconParametresWindow : FluentWindow
    {
        public BesanconParametresViewModel ViewModel { get; }

        public BesanconParametresWindow() : this(new BesanconParametresViewModel()) { }

        public BesanconParametresWindow(BesanconParametresViewModel vm)
        {
            InitializeComponent();
            ViewModel = vm;
            DataContext = ViewModel;
            ViewModel.CloseAction = result => { DialogResult = result; Close(); };
        }
    }
}
