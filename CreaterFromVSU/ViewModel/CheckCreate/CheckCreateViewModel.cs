using CreaterFromVSU.Constants;
using CreaterFromVSU.ViewModel.Utilites;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;

namespace CreaterFromVSU.ViewModel.CheckCreate
{
    class CheckCreateViewModel
    {
        public CheckCreateViewModel()
        {
        }
        public bool City_dist { get; set; }
        public bool City_ochno { get; set; }
        public bool Diplom_dist { get; set; }
        public bool Diplom_ochno { get; set; }
        public bool Sertific_dist { get; set; }
        public bool SertificFrom_dist { get; set; }
        public bool SertificFrom_ochno { get; set; }
        public bool Moder_dist { get; set; }
        public bool Moder_ochno { get; set; }
        public ICommand CheckCreateFilesCommand => new RelayCommand(CheckCreateFiles);
        private static void CheckCreateFiles()
        {
           /* if (!string.IsNullOrEmpty((Const.filePathOpen) && !string.IsNullOrEmpty(folderPathSave))
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
            }*/
        }
    }
}
