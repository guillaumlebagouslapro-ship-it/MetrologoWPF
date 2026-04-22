using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class DiagnosticGpibWindow : FluentWindow
    {
        public DiagnosticGpibWindow()
        {
            InitializeComponent();
            DataContext = new DiagnosticGpibViewModel();
        }
    }
}
