using System.Windows;
using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class CheminsStockageWindow : FluentWindow
    {
        public CheminsStockageWindow(CheminsStockageViewModel vm)
        {
            InitializeComponent();
            DataContext = vm;
            vm.CloseAction = ok =>
            {
                if (IsVisible) { DialogResult = ok; Close(); }
            };
        }
    }
}
