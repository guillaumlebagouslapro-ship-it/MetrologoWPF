using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class ParamsIncertWindow : FluentWindow
    {
        public ParamsIncertViewModel ViewModel { get; }

        public ParamsIncertWindow() : this(new ParamsIncertViewModel()) { }

        public ParamsIncertWindow(ParamsIncertViewModel vm)
        {
            InitializeComponent();
            ViewModel = vm;
            DataContext = ViewModel;
            ViewModel.CloseAction = result => { DialogResult = result; Close(); };
        }
    }
}
