using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace CreaterFromVSU.ViewModel.Utilites
{
    public abstract class BasicViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    }
}
