using System.Windows;

namespace Laba.Helpers
{
    public class UserHelper
    {
        /// <summary>
        /// Метод вывод пользовотелю диалоговое окно
        /// </summary>
        /// <param name="message"> Сообщение для пользователя </param>
        /// <param name="caption"> Наименование окна </param>
        /// <param name="messageBoxImage"> Задает значок, который отображается в окне сообщения </param>
        /// <returns> true - если пользователь согласен </returns>
        public static bool GetUserConfirmation(string message, string caption, MessageBoxImage messageBoxImage)
        {
            return MessageBox.Show(message, caption,
                    MessageBoxButton.YesNo,
                    messageBoxImage) == MessageBoxResult.Yes;
        }
    }
}
