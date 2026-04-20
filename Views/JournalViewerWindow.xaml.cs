using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class JournalViewerWindow : FluentWindow
    {
        public JournalViewerViewModel ViewModel { get; }

        public JournalViewerWindow()
        {
            InitializeComponent();
            ViewModel = new JournalViewerViewModel();
            DataContext = ViewModel;
        }
    }
}
