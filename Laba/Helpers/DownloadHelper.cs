using System.IO;
using System.Net;
using System.Windows;

namespace Laba.Helpers
{
    public class DownloadHelper
    {
        /// <summary>
        /// Метод загружает файл с исходными данными для локальной базы
        /// </summary>
        /// <returns> true - если файл успешно загружен </returns>
        public static bool DownloadFile()
        {
            if (File.Exists(@"security_threat_list.xlsx"))
            {
                if (!UserHelper.GetUserConfirmation("Файл с названием security_threat_list.xlsx уже существует. Перезаписать его?", 
                    "Загрузка файла security_threat_list.xlsx", 
                    MessageBoxImage.Question))
                {
                    return false;
                }
            }

            WebClient client = new WebClient();

            while (true)
            {
                try
                {
                    client.DownloadFile(@"https://bdu.fstec.ru/files/documents/thrlist.xlsx", @"security_threat_list.xlsx");
                    if (File.Exists(@"security_threat_list.xlsx"))
                    {
                        return true;
                    }
                }
                catch (WebException)
                {
                    if (!UserHelper.GetUserConfirmation("Нет соединения с Интернетом или открыт файл security_threat_list.xlsx. Попробовать снова?", 
                        "Ошибка при загрузке", 
                        MessageBoxImage.Error))
                    {
                        return false;
                    }
                }
            }
        }
    }
}
