using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class GestionAdressesLegacyWindow : FluentWindow
    {
        public GestionAdressesLegacyWindow(GestionAdressesLegacyViewModel vm)
        {
            InitializeComponent();
            DataContext = vm;
            vm.CloseAction = () =>
            {
                if (IsVisible) Close();
            };
        }
    }
}
