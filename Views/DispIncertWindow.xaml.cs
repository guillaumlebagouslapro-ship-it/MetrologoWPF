using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class DispIncertWindow : FluentWindow
    {
        public DispIncertViewModel ViewModel { get; }

        public DispIncertWindow()
        {
            InitializeComponent();
            ViewModel = new DispIncertViewModel();
            DataContext = ViewModel;
            ViewModel.CloseAction = _ => Close();
        }
    }
}
