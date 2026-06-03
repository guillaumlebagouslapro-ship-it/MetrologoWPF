using Metrologo.ViewModels;
using Wpf.Ui.Controls;

namespace Metrologo.Views
{
    public partial class JournalAdminWindow : FluentWindow
    {
        public JournalAdminViewModel ViewModel { get; }

        public JournalAdminWindow()
        {
            InitializeComponent();
            ViewModel = new JournalAdminViewModel();
            DataContext = ViewModel;
        }
    }
}
