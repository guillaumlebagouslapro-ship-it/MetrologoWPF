using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class GestionCatalogueRubidiumsWindow : FluentWindow
    {
        public GestionCatalogueRubidiumsViewModel ViewModel { get; }

        public GestionCatalogueRubidiumsWindow()
        {
            InitializeComponent();
            ViewModel = new GestionCatalogueRubidiumsViewModel();
            DataContext = ViewModel;
            ViewModel.CloseAction = result => { DialogResult = result; Close(); };
        }
    }
}
