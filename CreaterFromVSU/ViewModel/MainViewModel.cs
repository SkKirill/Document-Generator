using CreaterFromVSU.ViewModel.CheckCreate;
using CreaterFromVSU.ViewModel.Utilites;
using CreaterFromVSU.ViewModel.WorkConsole;
using Microsoft.VisualBasic.ApplicationServices;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;

namespace CreaterFromVSU.ViewModel
{
    class MainViewModel : BasicViewModel
    {
        public CheckCreateView CheckCreateView { get; set; }
        public WindowHelpInfo WindowHelpInfo { get; set; }
        public WorkConsoleView WorkConsoleView { get; set; }
        public string ImagePath
        {
            get => _imagePath;
            set
            {
                _imagePath = value;
                OnPropertyChanged("ImagePath");
            }
        }
        public string FolderPath
        {
            get => _folderPath;
            set
            {
                _folderPath = value;
                OnPropertyChanged("FolderPath");
            }
        }
        public string FileDataPath
        {
            get => _fileDataPath;
            set
            {
                _fileDataPath = value;
                OnPropertyChanged("FileDataPath");
            }
        }
        public ICommand OpenCheckCreateViewCommand => new RelayCommand(OpenCheckCreateView);
        public ICommand StartCreateComand => new RelayCommand(StartCreate);
        public ICommand OpenFolderCommand => new RelayCommand(OpenFolder);
        public ICommand BInfoOpenCommand => new RelayCommand(BInfoOpen);
        public ICommand OpenFileCommand => new RelayCommand(OpenFile);
        public ICommand OpenImageCommand => new RelayCommand(OpenImage);
        
        private string _fileDataPath;
        private string _folderPath;
        private string _imagePath;
        public MainViewModel()
        {

        }
        private void StartCreate()
        {
            WorkConsoleView = new WorkConsoleView();
            WorkConsoleView.Show();
        }
        private void OpenImage()
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                FileName = "Image",
                DefaultExt = ".png",
                Filter = "Image files (*.png;*.jpg;*.jpeg;*.gif;*.bmp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp"
            };
            dialog.ShowDialog();
        }
        private void OpenCheckCreateView()
        {
            CheckCreateView = new CheckCreateView();
            CheckCreateView.Show();
        }
        private void OpenFile()
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void OpenFolder()
        {
            try
            {
                CommonOpenFileDialog ofd = new() { IsFolderPicker = true };
                ofd.ShowDialog();
                if (!string.IsNullOrEmpty(ofd.FileName))
                {
                    /**/
                }
            }
            catch { }
        }
        private void BInfoOpen()
        {
            WindowHelpInfo windowHelp = new WindowHelpInfo();
            windowHelp.Show();
        }
    }
}
