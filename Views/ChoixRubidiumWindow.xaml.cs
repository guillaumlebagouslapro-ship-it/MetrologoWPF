using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class ChoixRubidiumWindow : FluentWindow
    {
        public ChoixRubidiumViewModel ViewModel { get; }

        public ChoixRubidiumWindow()
        {
            InitializeComponent();
            ViewModel = new ChoixRubidiumViewModel();
            DataContext = ViewModel;
            ViewModel.CloseAction = result => { DialogResult = result; Close(); };
        }
    }
}
