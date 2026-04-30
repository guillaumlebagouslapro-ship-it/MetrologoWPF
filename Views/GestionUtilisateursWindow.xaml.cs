using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class GestionUtilisateursWindow : FluentWindow
    {
        public GestionUtilisateursWindow(GestionUtilisateursViewModel vm)
        {
            InitializeComponent();
            DataContext = vm;
        }
    }
}
