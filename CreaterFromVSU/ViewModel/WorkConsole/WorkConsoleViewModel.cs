using CreaterFromVSU.ViewModel.Utilites;

namespace CreaterFromVSU.ViewModel.WorkConsole
{
    class WorkConsoleViewModel : BasicViewModel
    {
        private string _logText;

        public WorkConsoleViewModel()
        {
        }

        public string LogText {  
            get 
            { 
                return _logText; 
            } 
            set 
            { 
                _logText = value;
                OnPropertyChanged("LogText");
            } 
        }
    }
}
