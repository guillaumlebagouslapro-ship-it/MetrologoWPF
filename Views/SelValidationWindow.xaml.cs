using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class SelValidationWindow : FluentWindow
    {
        public SelValidationViewModel ViewModel { get; }

        public SelValidationWindow() : this(new SelValidationViewModel()) { }

        public SelValidationWindow(SelValidationViewModel vm)
        {
            InitializeComponent();
            ViewModel = vm;
            DataContext = ViewModel;
            ViewModel.CloseAction = result => { DialogResult = result; Close(); };
        }
    }
}
