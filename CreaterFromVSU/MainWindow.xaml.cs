using Microsoft.Win32;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MyWinForm = System.Windows.Forms;
using MyWinApiPack = Microsoft.WindowsAPICodePack;
 

namespace CreaterFromVSU
{
    public partial class MainWindow : Window
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

        private void ButtonCollapse_MouseLeave(object sender, MouseEventArgs e)
        {
            ButtonCollapse.Background = new SolidColorBrush(Color.FromRgb(62, 62, 246));
        }

        private void ButtonCollapse_MouseEnter(object sender, MouseEventArgs e)
        {
            ButtonCollapse.Background = new SolidColorBrush(Color.FromRgb(121, 121, 227));
        }

        private void ButtonExsit_MouseLeave(object sender, MouseEventArgs e)
        {
            ButtonExsit.Background = new SolidColorBrush(Color.FromRgb(255, 0, 0));
        }

        private void ButtonExsit_MouseEnter(object sender, MouseEventArgs e)
        {
            ButtonExsit.Background = new SolidColorBrush(Color.FromRgb(255, 80, 80));
        }

        private void TopPanel_MouseUp(object sender, MouseButtonEventArgs e)
        {
            isDragging = false;
        }

        private void Image_MouseDown(object sender, MouseButtonEventArgs e)
        {
            Environment.Exit(0);
        }

        private void TopPanel_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                Point currentPosition = e.GetPosition(this);
                Left = Left - (lastPosition.X - currentPosition.X);
                Top = Top - (lastPosition.Y - currentPosition.Y);
            }
        }

        private void TopPanel_MouseDown(object sender, MouseButtonEventArgs e)
        {
            isDragging = true;
            lastPosition = e.GetPosition(this);
        }

        private void ButtonInfoOpen_MouseEnter(object sender, MouseEventArgs e)
        {
            ButtonInfoOpen.Background = new SolidColorBrush(Color.FromRgb(153, 153, 153));
        }

        private void buttonOpenFileExcel_MouseEnter(object sender, MouseEventArgs e)
        {
            buttonOpenFileExcel.Background = new SolidColorBrush(Color.FromRgb(153, 153, 153));
        }

        private void buttonOpenFolder_MouseEnter(object sender, MouseEventArgs e)
        {
            buttonOpenFolder.Background = new SolidColorBrush(Color.FromRgb(153, 153, 153));
        }
        private void buttonOpenFilePodl_MouseEnter(object sender, MouseEventArgs e)
        {
            buttonOpenFilePodl.Background = new SolidColorBrush(Color.FromRgb(153, 153, 153));
        }

        private void buttonStartProgram_MouseEnter(object sender, MouseEventArgs e)
        {
            buttonStartProgram.Background = new SolidColorBrush(Color.FromRgb(153, 153, 153));
        }

        private void ButtonInfoOpen_MouseLeave(object sender, MouseEventArgs e)
        {
            ButtonInfoOpen.Background = new SolidColorBrush(Color.FromRgb(73, 73, 73));
        }

        private void buttonStartProgram_MouseLeave(object sender, MouseEventArgs e)
        {
            buttonStartProgram.Background = new SolidColorBrush(Color.FromRgb(73, 73, 73));
        }

        private void buttonOpenFolder_MouseLeave(object sender, MouseEventArgs e)
        {
            buttonOpenFolder.Background = new SolidColorBrush(Color.FromRgb(73, 73, 73));
        }
        private void buttonOpenFilePodl_MouseLeave(object sender, MouseEventArgs e)
        {
            buttonOpenFilePodl.Background = new SolidColorBrush(Color.FromRgb(73, 73, 73));
        }

        private void buttonOpenFileExcel_MouseLeave(object sender, MouseEventArgs e)
        {
            buttonOpenFileExcel.Background = new SolidColorBrush(Color.FromRgb(73, 73, 73));
        }

        private void Image_MouseDown_1(object sender, MouseButtonEventArgs e)
        {
            WindowState = WindowState.Minimized;
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
        private void buttonOpenFileExcel_MouseDown(object sender, MouseButtonEventArgs e)
        {
            filePathOpen = openFileDialog();
            LableFileExcel.Content = filePathOpen;
        }
        private void buttonOpenFilePodl_MouseDown(object sender, MouseButtonEventArgs e)
        {
            filePathPicture = openFileDialog();
            LablePodlFile.Content = filePathPicture;
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
                    LableFolder.Content = folderPathSave;
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
                        MessageBox.Show(ERROR_MESSAGE_PATH_TABLE, "Ошибка.");
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