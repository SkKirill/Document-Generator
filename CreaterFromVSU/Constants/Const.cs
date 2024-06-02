using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreaterFromVSU.Constants
{
    public class Const
    {
        public const string ERROR_MESSAGE_PATH_PICTURE = @"Вы не указали расположение файла подложки(картинка/скан сертификата). Попробуйте еще раз!";
        public const string ERROR_MESSAGE_TITLE = @"Ошибка!";
        public const string ERROR_MESSAGE_PATH_TABLE = @"Вы не указали расположение файла таблицы!. Попробуйте еще раз!";
        public const string ERROR_MESSAGE_CHECK_PATH = @"При обработке расположений файлов произошла ошибка. Проверьте и попробуйте еще раз!";

        // Сообщение о создании прогой дефолтных значений
        public const string WARNING_MESSAGE_TITLE = @"Проверьте!";
        public const string WARNING_MESSAGE_CREATE_FOLDER_SAVE = @"Путь для соранения будет выбран автомотически: { }";

        public string filePathOpen;
        public string folderPathSave;
        public string filePathPicture;
    }
}
