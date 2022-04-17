using System.Collections.Generic;
using System.Text;
using System.Windows;

namespace Laba
{
    /// <summary>
    /// Логика взаимодействия для RefreshInformationWindow.xaml
    /// </summary>
    public partial class RefreshInformationWindow : Window
    {
        public RefreshInformationWindow(List<string> added, List<string> changed, List<string> deleted, int updatedCount)
        {
            InitializeComponent();
            MessageTextBlock.Text = "Локальная база данных успешно обновлена!\n" +
                $"Количество обновленных записей: {updatedCount}\n\n";

            var sb = new StringBuilder("Добавлено:\n");

            foreach (var item in added)
            {
                sb.AppendLine(item);
            }

            sb.AppendLine();
            sb.AppendLine("Изменено:");

            foreach (var item in changed)
            {
                sb.AppendLine(item);
            }

            sb.AppendLine();
            sb.AppendLine("Удалено:");

            foreach (var item in deleted)
            {
                sb.AppendLine(item);
            }

            MessageTextBlock.Text += sb.ToString();
        }

        /// <summary>
        /// "OK"
        /// </summary>
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}
