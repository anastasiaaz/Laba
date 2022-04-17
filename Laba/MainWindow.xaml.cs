using Laba.Enums;
using Laba.Helpers;
using Laba.Model;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

namespace Laba
{
    public partial class MainWindow : Window
    {
        private Thread autoRefresh = new Thread(DataBaseHelper.AutoRefreshLocalDataBase);

        private Pagination pagination;

        public MainWindow()
        {
            pagination = new Pagination();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            InitializeComponent();

            DataBaseHelper.CheckLocalDataBase();

            autoRefresh.Start();
        }

       /// <summary>
       /// "СОХРАНИТЬ"
       /// </summary>
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (!DataBaseHelper.ValidateLocalDataBase())
            {
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.DefaultExt = ".xlsx";
            saveFileDialog.FileName = "security_threat_list";

            if (saveFileDialog.ShowDialog() == false)
            {
                return;
            }

            var file = new FileInfo(@"LocalDataBase.xlsx");
            using (var package = new ExcelPackage(file))
            {
                var sheet = package.Workbook.Worksheets.First();

                foreach (var cell in sheet.Cells)
                {
                    if (cell.Address.Contains("F") || cell.Address.Contains("G") || cell.Address.Contains("H"))
                    {
                        if (!ParseBoolToString(sheet, cell.Address.Substring(0, 1), int.Parse(cell.Address.Substring(1))))
                        {
                            MessageBox.Show("Похоже что-то не так с базой. Попробуйте ее обновить.");
                            return;
                        }
                    }
                }

                sheet.InsertRow(1, 1);
                sheet.Cells["A1"].Value = "Идентификатор УБИ";
                sheet.Cells["B1"].Value = "Наименование УБИ";
                sheet.Cells["C1"].Value = "Описание";
                sheet.Cells["D1"].Value = "Источник угрозы";
                sheet.Cells["E1"].Value = "Объект воздействия";
                sheet.Cells["F1"].Value = "Нарушение конфиденциальности";
                sheet.Cells["G1"].Value = "Нарушение целостности";
                sheet.Cells["H1"].Value = "Нарушение доступности";

                package.SaveAs(saveFileDialog.FileName);
                MessageBox.Show("Успешно сохранено!");
            }
        }

        /// <summary>
        /// "ПРОСМОТР СВЕДЕНИЙ ОБ УГРОЗЕ"
        /// </summary>
        private void ThreatInformationButton_Click(object sender, RoutedEventArgs e)
        {
            if (!DataBaseHelper.ValidateLocalDataBase())
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(IDTextBox.Text) || string.IsNullOrEmpty(IDTextBox.Text))
            {
                MessageBox.Show("Идентификатор не введен!");
                return;
            }

            var file = new FileInfo(@"LocalDataBase.xlsx");
            using (var package = new ExcelPackage(file))
            {
                var sheet = package.Workbook.Worksheets.First();

                var cellID = sheet.Cells.FirstOrDefault(cell => cell.Address.Contains("A") && cell.Value.ToString() == IDTextBox.Text.Trim());

                if (cellID == null)
                {
                    MessageBox.Show($"Запись с идентификатором {IDTextBox.Text} не найдена.");
                    return;
                }

                var rowNumber = int.Parse(cellID.Address.Replace("A", ""));

                if (!ParseBoolToString(sheet, "F", rowNumber))
                {
                    MessageBox.Show("Похоже что-то не так с базой. Попробуйте ее обновить.");
                    return;
                }

                if (!ParseBoolToString(sheet, "G", rowNumber))
                {
                    MessageBox.Show("Похоже что-то не так с базой. Попробуйте ее обновить.");
                    return;
                }

                if (!ParseBoolToString(sheet, "H", rowNumber))
                {
                    MessageBox.Show("Похоже что-то не так с базой. Попробуйте ее обновить.");
                    return;
                }

                TextBlock.Text = $"Идентификатор УБИ: {cellID.Value}.\n\n" +
                $"Наименование УБИ: {sheet.Cells[$"B{rowNumber}"].Value}.\n\n" +
                $"Описание: {sheet.Cells[$"C{rowNumber}"].Value}.\n\n" +
                $"Источник угрозы: {sheet.Cells[$"D{rowNumber}"].Value}.\n\n" +
                $"Объект воздействия: {sheet.Cells[$"E{rowNumber}"].Value}.\n\n" +
                $"Нарушение конфиденциальности: {sheet.Cells[$"F{rowNumber}"].Value}.\n\n" +
                $"Нарушение целостности: {sheet.Cells[$"G{rowNumber}"].Value}.\n\n" +
                $"Нарушение доступности: {sheet.Cells[$"H{rowNumber}"].Value}.";

                TextBlock.Visibility = Visibility.Visible;

                DataGrid.Visibility = Visibility.Hidden;                
                NextPageButton.Visibility = Visibility.Hidden;
                PreviousPageButton.Visibility = Visibility.Hidden;
            }
        }

        /// <summary>
        /// Метод для проверки введенного значения в строку Идентификатора
        /// </summary>
        private void IDTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = !char.IsDigit(e.Text[0]);
        }

        /// <summary>
        /// Метод для замены "0" и "1" на "да" и "нет"
        /// </summary>
        /// <param name="sheet"> Excel таблица с базой данных </param>
        /// <param name="letter"> Значение столбца </param>
        /// <param name="rowNumber"> Значение строки </param>
        /// <returns> true - если изменения успешно внесены </returns>
        private static bool ParseBoolToString(ExcelWorksheet sheet, string letter, int rowNumber)
        {
            if (int.TryParse(sheet.Cells[$"{letter}{rowNumber}"].Value.ToString(), out int boolValue))
            {
                if (boolValue == 1)
                {
                    sheet.Cells[$"{letter}{rowNumber}"].Value = "Да";
                    return true;
                }
                else if (boolValue == 0)
                {
                    sheet.Cells[$"{letter}{rowNumber}"].Value = "Нет";
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return false;
        }

        /// <summary>
        /// "ПРОСМОТР УГРОЗ" - список
        /// </summary>
        private void ThreatViewingButton_Click(object sender, RoutedEventArgs e)
        {
            if (!DataBaseHelper.ValidateLocalDataBase())
            {
                return;
            }

            pagination.Data.Clear();

            var file = new FileInfo(@"LocalDataBase.xlsx");
            using (var package = new ExcelPackage(file))
            {
                var sheet = package.Workbook.Worksheets.First();

                foreach (var row in sheet.Cells.Where(cell => cell.Address.Contains("A") || cell.Address.Contains("B")).GroupBy(cell => cell.Address.Substring(1)))
                {
                    pagination.Data.Add(new ShortNote(row.First().Text, row.Last().Text));
                }
            }

            pagination.PageIndex = 1;
            DataGrid.ItemsSource = pagination.Data.Take(Pagination.numberOfRecPerPage);

            DataGrid.Visibility = Visibility.Visible;
            NextPageButton.Visibility = Visibility.Visible;
            PreviousPageButton.Visibility = Visibility.Visible;
            PreviousPageButton.IsEnabled = false;

            TextBlock.Visibility = Visibility.Hidden;

            DataGrid.Columns[0].Header = "Идентификатор УБИ";
            DataGrid.Columns[1].Header = "Наименование УБИ";
            DataGrid.Columns[0].Width = 150;
            DataGrid.Columns[1].Width = 1400;
        }

        /// <summary>
        /// Метод для корректной работы кнопок "Вперед" и "Назад"
        /// </summary>
        /// <param name="mode"> Направление смены страниц </param>
        private void Navigate(PagingMode mode)
        {
            switch (mode)
            {
                case PagingMode.Next:
                    PreviousPageButton.IsEnabled = true;

                    if (pagination.Data.Skip(pagination.PageIndex *
                    Pagination.numberOfRecPerPage).Take(Pagination.numberOfRecPerPage).Count() == 0)
                    {
                        DataGrid.ItemsSource = null;
                        DataGrid.ItemsSource = pagination.Data.Skip((pagination.PageIndex *
                        Pagination.numberOfRecPerPage) - Pagination.numberOfRecPerPage).Take(Pagination.numberOfRecPerPage);
                    }
                    else
                    {
                        DataGrid.ItemsSource = null;
                        DataGrid.ItemsSource = pagination.Data.Skip(pagination.PageIndex *
                        Pagination.numberOfRecPerPage).Take(Pagination.numberOfRecPerPage);
                        pagination.PageIndex++;
                    }

                    if (pagination.Data.Count <= (pagination.PageIndex * Pagination.numberOfRecPerPage))
                    {
                        NextPageButton.IsEnabled = false;
                    }

                    break;

                case PagingMode.Previous:

                    NextPageButton.IsEnabled = true;

                    pagination.PageIndex--;
                    DataGrid.ItemsSource = null;

                    DataGrid.ItemsSource = pagination.Data.Skip
                    ((pagination.PageIndex - 1) * Pagination.numberOfRecPerPage).Take(Pagination.numberOfRecPerPage);

                    if (pagination.PageIndex == 1)
                    {
                        PreviousPageButton.IsEnabled = false;
                    }

                    break;
            }
        }

        /// <summary>
        /// Кнопка "Назад"
        /// </summary>
        private void PreviousPageButton_Click(object sender, RoutedEventArgs e)
        {
            Navigate(PagingMode.Previous);
            DataGrid.Columns[0].Header = "Идентификатор УБИ";
            DataGrid.Columns[1].Header = "Наименование УБИ";
            DataGrid.Columns[0].Width = 150;
            DataGrid.Columns[1].Width = 1400;
        }

        /// <summary>
        /// Кнопка "Вперед"
        /// </summary>
        private void NextPageButton_Click(object sender, RoutedEventArgs e)
        {
            Navigate(PagingMode.Next);
            DataGrid.Columns[0].Header = "Идентификатор УБИ";
            DataGrid.Columns[1].Header = "Наименование УБИ";
            DataGrid.Columns[0].Width = 150;
            DataGrid.Columns[1].Width = 1400;
        }

        /// <summary>
        /// Метод прерывания потока автообновления по закрытию окна программы
        /// </summary>
        private void Window_Closed(object sender, EventArgs e)
        {
            autoRefresh.Abort();
        }

        /// <summary>
        /// "ОБНОВИТЬ"
        /// </summary>
        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            DataBaseHelper.RefreshLocalDataBase();
        }
    }
}
