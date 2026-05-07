using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class GestionModulesIncertitudeWindow : FluentWindow
    {
        public GestionModulesIncertitudeWindow()
        {
            InitializeComponent();
            DataContext = new GestionModulesIncertitudeViewModel();
        }
    }
}
