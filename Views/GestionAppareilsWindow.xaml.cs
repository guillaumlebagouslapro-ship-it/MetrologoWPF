using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class GestionAppareilsWindow : FluentWindow
    {
        public GestionAppareilsWindow(GestionAppareilsViewModel vm)
        {
            InitializeComponent();
            DataContext = vm;
        }
    }
}
