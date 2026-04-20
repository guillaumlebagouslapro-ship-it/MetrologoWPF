using Metrologo.Models;
using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class SelectionGateWindow : FluentWindow
    {
        public SelectionGateViewModel ViewModel { get; }

        public SelectionGateWindow(Mesure mesure)
        {
            InitializeComponent();
            ViewModel = new SelectionGateViewModel(mesure);
            DataContext = ViewModel;
            ViewModel.CloseAction = result => { DialogResult = result; Close(); };
        }
    }
}
