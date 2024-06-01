using Microsoft.Win32;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using MyWinForm = System.Windows.Forms;
using MyWinApiPack = Microsoft.WindowsAPICodePack;
using MahApps.Metro.Controls;


namespace CreaterFromVSU
{
    public partial class MainWindow : MetroWindow
    {
        // Сообщения об ошибки пользователю //
        private const string ERROR_MESSAGE_PATH_PICTURE = @"Вы не указали расположение файла подложки(картинка/скан сертификата). Попробуйте еще раз!";
        private const string ERROR_MESSAGE_TITLE = @"Ошибка!";
        private const string ERROR_MESSAGE_PATH_TABLE = @"Вы не указали расположение файла таблицы!. Попробуйте еще раз!";
        private const string ERROR_MESSAGE_CHECK_PATH = @"При обработке расположений файлов произошла ошибка. Проверьте и попробуйте еще раз!";
        
        // Сообщение о создании прогой дефолтных значений
        private const string WARNING_MESSAGE_TITLE = @"Проверьте!";
        private const string WARNING_MESSAGE_CREATE_FOLDER_SAVE = @"Путь для соранения будет выбран автомотически: { }";

        
        // Управление расположением окна при нажатии на borderTopPanel //
        private bool isDragging = false;
        private Point lastPosition;

        // Пути к кфайлам //
        private string filePathOpen;
        private string folderPathSave;
        private string filePathPicture;

        public MainWindow()
        {
            filePathOpen = string.Empty;
            folderPathSave = string.Empty;
            filePathPicture = string.Empty;
            InitializeComponent();
        }

        private string openFileDialog()
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == true)
                {
                    return openFileDialog.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return string.Empty;
        }
      
        private void buttonOpenFolder_MouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                MyWinApiPack.Dialogs.CommonOpenFileDialog ofd = new() { IsFolderPicker = true };
                ofd.ShowDialog();
                if (!string.IsNullOrEmpty(ofd.FileName))
                {
                    folderPathSave = ofd.FileName;
                }
            }
            catch { }
        }

        private void ButtonInfoOpen_MouseDown(object sender, MouseButtonEventArgs e)
        {
            WindowHelp windowHelp = new WindowHelp();
            windowHelp.Show();
        }

        private void buttonStartProgram_MouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Window1 window1 = new Window1();
                window1.ShowDialog();
                if (window1.start)
                {
                    if (string.IsNullOrEmpty(folderPathSave))
                    {
                        string? path = System.Reflection.Assembly.GetExecutingAssembly().Location;
                        string? directory = System.IO.Path.GetDirectoryName(path);

                        MyWinForm.DialogResult result = MyWinForm.MessageBox.Show(string.Format(WARNING_MESSAGE_CREATE_FOLDER_SAVE, directory),
                            WARNING_MESSAGE_TITLE, MyWinForm.MessageBoxButtons.YesNo);
                        if (result.Equals(MyWinForm.DialogResult.Yes))
                        {
                           /* folderPathSave = directory is null ? string.Empty : directory;*/
                        }
                        if (string.IsNullOrEmpty(folderPathSave))
                        {
                            throw new Exception();
                        }
                    }

                    if (!string.IsNullOrEmpty(filePathOpen) && !string.IsNullOrEmpty(folderPathSave))
                    {
                        WindowLog.Create_for is_cr = new WindowLog.Create_for();
                        is_cr.city_dist = (bool)window1.city_dist.IsChecked;
                        is_cr.city_ochno = (bool)window1.city_ochno.IsChecked;
                        is_cr.diplom_dist = (bool)window1.diplom_dist.IsChecked;
                        is_cr.diplom_ochno = (bool)window1.diplom_ochno.IsChecked;
                        is_cr.sertific_dist = (bool)window1.sertific_dist.IsChecked;
                        is_cr.sertific_ochno = (bool)window1.sertific_ochno.IsChecked;
                        is_cr.sertificFrom_dist = (bool)window1.sertificFrom_dist.IsChecked;
                        is_cr.sertificFrom_ochno = (bool)window1.sertificFrom_ochno.IsChecked;
                        is_cr.moder_dist = (bool)window1.moder_dist.IsChecked;
                        is_cr.moder_ochno = (bool)window1.moder_ochno.IsChecked;
                        if (is_cr.sertificFrom_dist || is_cr.sertificFrom_ochno)
                        {
                            if (string.IsNullOrEmpty(filePathPicture))
                            {
                                MessageBox.Show(ERROR_MESSAGE_PATH_PICTURE, ERROR_MESSAGE_TITLE);
                                return;
                            }
                        }
                        WindowLog windowLog = new WindowLog(is_cr, filePathOpen, filePathPicture, folderPathSave);
                        windowLog.Show();
                    }
                    else
                    {
                        MessageBox.Show(ERROR_MESSAGE_PATH_TABLE, "Ошибка!");
                    }
                }
            }
            catch (Exception es)
            {
                MessageBox.Show(ERROR_MESSAGE_CHECK_PATH + "\n" + es, ERROR_MESSAGE_TITLE);
            }
        }
    }
}