using System.Windows.Input;

namespace CreaterFromVSU.ViewModel.Utilites
{
    public class RelayCommand : ICommand
    {
        private Action _methodToExecute;
        private Func<bool> _canExecuteMethod;

        public RelayCommand(Action methodToExecute, Func<bool> canExecuteMethod = null)
        {
            _methodToExecute = methodToExecute;
            _canExecuteMethod = canExecuteMethod;
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter)
        {
            return _canExecuteMethod == null || _canExecuteMethod();
        }

        public void Execute(object parameter)
        {
            _methodToExecute?.Invoke();
        }
    }
}
