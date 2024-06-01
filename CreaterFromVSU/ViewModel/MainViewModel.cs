using CreaterFromVSU.ViewModel.CheckCreate;
using CreaterFromVSU.ViewModel.Utilites;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace CreaterFromVSU.ViewModel
{
    class MainViewModel : BasicViewModel
    {
        public CheckCreateView CheckCreateView { get; set; }
        public WindowHelpInfo WindowHelpInfo { get; set; }
        public MainViewModel()
        {

        }

        public ICommand OpenCheckCreateViewCommand => new RelayCommand(OpenCheckCreateView);

        private void OpenCheckCreateView()
        {
            CheckCreateView = new CheckCreateView();
            CheckCreateView.Show();
        }
        public ICommand OpenCheckHelpInfoViewCommand => new RelayCommand(OpenCheckHelpInfoView);

        private void OpenCheckHelpInfoView()
        {
            WindowHelpInfo = new WindowHelpInfo();
            WindowHelpInfo.Show();
        }
    }
}
