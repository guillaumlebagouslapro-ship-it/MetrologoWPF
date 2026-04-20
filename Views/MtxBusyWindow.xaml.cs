using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class MtxBusyWindow : FluentWindow
    {
        public MtxBusyViewModel ViewModel { get; }

        public MtxBusyWindow()
        {
            InitializeComponent();
            ViewModel = new MtxBusyViewModel();
            DataContext = ViewModel;
            ViewModel.CloseAction = _ => Close();
        }
    }
}
