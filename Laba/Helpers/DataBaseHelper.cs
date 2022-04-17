using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;

namespace Laba.Helpers
{
    public class DataBaseHelper
    {
        private static List<string> cellsAdresses = new List<string>() { "B", "C", "D", "E", "F", "G", "H" };

        private static Dictionary<int, char> allCellsAdresses = new Dictionary<int, char>() 
        { 
            { 1, 'A' },
            { 2, 'B' },
            { 3, 'C' },
            { 4, 'D' },
            { 5, 'E' },
            { 6, 'F' },
            { 7, 'G' },
            { 8, 'H' } 
        };

        /// <summary> 
        /// Метод проверяет наличие файла локальной базы 
        /// </summary>
        /// <returns> true - если файл локальной базы найден </returns>
        public static bool CheckLocalDataBase()
        {
            if (!File.Exists(@"LocalDataBase.xlsx"))
            {
                if (!UserHelper.GetUserConfirmation("Файл локальной базы не найден. Создать его?",
                    "Ошибка \"Файл не найден\"",
                    MessageBoxImage.Error))
                {
                    return false;
                }

                UpdateLocalDataBase("Локальная база успешно создана.");
            }
            return true;

        }

        /// <summary>
        /// Метод обновления локальной базы
        /// </summary>
        /// <param name="message"> Сообщение при успешном создании </param>
        private static void UpdateLocalDataBase(string message)
        {
            bool infofile = File.Exists(@"security_threat_list.xlsx");

            if (!DownloadHelper.DownloadFile())
            {
                return;
            }

            CreateLocalDataBase(message);

            if (!infofile)
            {
                File.Delete(@"security_threat_list.xlsx");
            }
        }

        /// <summary>
        /// Метод создания локальной базы
        /// </summary>
        /// <param name="message"> Сообщение при успешном создании </param>
        private static void CreateLocalDataBase(string message)
        {
            while (true)
            {
                try
                {
                    var file = new FileInfo(@"security_threat_list.xlsx");
                    using (var package = new ExcelPackage(file))
                    {
                        var sheet = package.Workbook.Worksheets.First();

                        sheet.DeleteRow(1, 2);
                        sheet.DeleteColumn(9, 10);

                        package.SaveAs("LocalDataBase.xlsx");
                    }
                    if (File.Exists(@"LocalDataBase.xlsx"))
                    {
                        MessageBox.Show(message,
                            "Файл создан");
                        return;
                    }
                }
                catch (InvalidOperationException)
                {
                    if (!UserHelper.GetUserConfirmation("Не удалось создать базу. Возможно открыт файл LocalDataBase.xlsx. Попробовать снова?",
                        "Ошибка создания базы",
                        MessageBoxImage.Error))
                    {
                        return;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ой-ой, что-то пошло не так :( Возможно, это поможет: {ex.Message}");
                    return;
                }
            }
        }

        /// <summary>
        /// Метод проверяет наличие пустых ячеек в базе
        /// </summary>
        /// <returns> true - если пустных ячеек не обнаружено </returns>
        public static bool ValidateLocalDataBase()
        {
            if (!CheckLocalDataBase())
            {
                return false;
            }

            var file = new FileInfo(@"LocalDataBase.xlsx");
            using (var package = new ExcelPackage(file))
            {
                var sheet = package.Workbook.Worksheets.First();
                
                var currentCell = 1;

                foreach (var cell in sheet.Cells)
                {
                    if (!cell.Address.Contains(allCellsAdresses[currentCell]))
                    {
                        MessageBox.Show("Похоже что-то не так с базой. Попробуйте ее обновить.");
                        return false;
                    }

                    currentCell++;
                    if (currentCell == 9)
                    {
                        currentCell = 1;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Метод автоматического обновления базы, запускается в отдельном потоке
        /// </summary>
        public static void AutoRefreshLocalDataBase()
        {
            while (true)
            {
                Thread.Sleep(43200 * 1000);
                MessageBox.Show("Начинается автоматическое обновление базы.");
                RefreshLocalDataBase();
            }
        }

        /// <summary>
        /// Метод обновляет локальную базу данных на основе оригинальной базы данных
        /// </summary>
        public static void RefreshLocalDataBase()
        {
            if (!File.Exists(@"LocalDataBase.xlsx"))
            {
                UpdateLocalDataBase("Файл локальной базы полностью обновлен.");
                return;
            }

            bool infofile = File.Exists(@"security_threat_list.xlsx");

            if (!DownloadHelper.DownloadFile())
            {
                return;
            }

            while (true)
            {
                var changed = new List<string>();
                var deleted = new List<string>();
                var added = new List<string>();
                var updatedCount = 0;

                try
                {
                    var file = new FileInfo(@"security_threat_list.xlsx");
                    using (var package = new ExcelPackage(file))
                    {
                        var sheet = package.Workbook.Worksheets.First();

                        sheet.DeleteRow(1, 2);
                        sheet.DeleteColumn(9, 10);

                        var localFile = new FileInfo(@"LocalDataBase.xlsx");
                        using (var packageLocal = new ExcelPackage(localFile))
                        {
                            var sheetLocal = packageLocal.Workbook.Worksheets.First();

                            foreach (var cellWithId in sheet.Cells.Where(cell => cell.Address.Contains("A")))
                            {
                                var rowNumber = int.Parse(cellWithId.Address.Replace("A", ""));
                                var localCell = sheetLocal.Cells.FirstOrDefault(cell => cell.Text == cellWithId.Text && cell.Address.Contains('A'));
                                if (localCell != null)
                                {
                                    var rowNumberLocal = int.Parse(localCell.Address.Replace("A", ""));

                                    var changes = GetRowChanges(sheet, sheetLocal, rowNumber, rowNumberLocal);

                                    if (changes.Any())
                                    {
                                        changed.Add($"УБИ.{cellWithId.Text}:");
                                        changed.AddRange(changes);

                                        updatedCount++;
                                    }                                    
                                }
                                else
                                {
                                    sheetLocal.InsertRow(rowNumber, 1);
                                    sheetLocal.Cells[cellWithId.Address].Value = cellWithId.Value;
                                    added.Add($"УБИ.{cellWithId.Text}:");

                                    added.AddRange(GetNewRow(sheet, sheetLocal, rowNumber));

                                    updatedCount++;
                                }
                            }

                            foreach (var cellWithId in sheetLocal.Cells.Where(cell => cell.Address.Contains("A")))
                            {
                                var rowNumber = int.Parse(cellWithId.Address.Replace("A", ""));

                                if (!int.TryParse(cellWithId.Text, out int _))
                                {
                                    deleted.Add($"Строчка с не чсиловым идентификатором \"{cellWithId.Text}\":");

                                    deleted.AddRange(GetRowInformation(sheetLocal, rowNumber));

                                    sheetLocal.DeleteRow(rowNumber);

                                    updatedCount++;
                                    continue;
                                }

                                var cell = sheet.Cells.FirstOrDefault(cellLocal => cellLocal.Text == cellWithId.Text);
                                if (cell == null)
                                {
                                    deleted.Add($"Лишняя строчка с идентификатором {cellWithId.Text}:");

                                    deleted.AddRange(GetRowInformation(sheetLocal, rowNumber));

                                    sheetLocal.DeleteRow(rowNumber);

                                    updatedCount++;
                                    continue;
                                }
                            }

                            var currentCell = 1;
                            var rowsToDelete = new HashSet<int>();

                            foreach (var cell in sheetLocal.Cells)
                            {
                                if (!cell.Address.Contains(allCellsAdresses[currentCell]) && allCellsAdresses[currentCell] == 'A')
                                {
                                    var rowNumber = int.Parse(cell.Address.Substring(1));

                                    rowsToDelete.Add(rowNumber);
                                    currentCell++;
                                }

                                currentCell++;
                                if (currentCell == 9)
                                {
                                    currentCell = 1;
                                }
                            }

                            foreach (var rowNumber in rowsToDelete.OrderByDescending(x => x))
                            {
                                deleted.Add($"Лишняя строчка без идентификатора:");

                                deleted.AddRange(GetRowInformation(sheetLocal, rowNumber));

                                sheetLocal.DeleteRow(rowNumber);

                                updatedCount++;
                            }

                            packageLocal.Save();
                        }
                    }

                    if (updatedCount == 0)
                    {
                        MessageBox.Show("Локальная база данных содержит актуальную информацию. Обновление не требовалось.");
                    }
                    else
                    {
                        RefreshInformationWindow refreshInformationWindow = new RefreshInformationWindow(added, changed, deleted, updatedCount);
                        refreshInformationWindow.ShowDialog();
                    }
                    break;
                }
                catch (InvalidOperationException)
                {
                    if (!UserHelper.GetUserConfirmation("Не удалось обновить базу. Возможно открыт файл LocalDataBase.xlsx. Попробовать снова?",
                        "Ошибка обновления базы",
                        MessageBoxImage.Error))
                    {
                        break;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ой-ой, что-то пошло не так :( Возможно, это поможет: {ex.Message}");
                    break;
                }
            }
            if (!infofile)
            {
                File.Delete(@"security_threat_list.xlsx");
            }
        }

        /// <summary>
        /// Метод обновляет строчку определенной угрозы в локальном файле из загруженного файла
        /// </summary>
        /// <param name="worksheet">Лист загруженного файла</param>
        /// <param name="worksheetLocal">Лист локальной базы</param>
        /// <param name="rowNumber">Номер строчки в загруженном файле</param>
        /// <param name="rowNumberLocal">Номер строчки в локальной базе</param>
        /// <returns>Список изменений</returns>
        private static List<string> GetRowChanges(ExcelWorksheet worksheet, ExcelWorksheet worksheetLocal, int rowNumber, int rowNumberLocal)
        {
            List<string> changes = new List<string>();

            foreach (var cellAddres in cellsAdresses)
            {
                var cell = worksheet.Cells[$"{cellAddres}{rowNumber}"];
                var cellLocal = worksheetLocal.Cells[$"{cellAddres}{rowNumberLocal}"];

                if (cellLocal.Text != cell.Text)
                {
                    changes.Add($"{cellLocal.Text} -> {cell.Text}");
                    cellLocal.Value = cell.Value;
                }
            }

            return changes;
        }

        /// <summary>
        /// Метод получает информацию из строк, подлежащих удалению в локальной базе
        /// </summary>
        /// <param name="worksheet">Лист локальной базы</param>
        /// <param name="rowNumber">Номер строчки в локальной базе</param>
        /// <returns>Список удаленных данных</returns>
        private static List<string> GetRowInformation(ExcelWorksheet worksheet, int rowNumber)
        {
            List<string> information = new List<string>();

            foreach (var cellAddres in cellsAdresses)
            {
                var cell = worksheet.Cells[$"{cellAddres}{rowNumber}"];

                if (!string.IsNullOrEmpty(cell.Text))
                {
                    information.Add(cell.Text);
                }                
            }

            return information;
        }

        /// <summary>
        /// Метод заполняет новую строку в локальном файле из загруженного файла
        /// </summary>
        /// <param name="worksheet">Лист загруженного файла</param>
        /// <param name="worksheetLocal">Лист локальной базы</param>
        /// <param name="rowNumber">Номер строчки в обоих файлах</param>
        /// <returns></returns>
        private static List<string> GetNewRow(ExcelWorksheet worksheet, ExcelWorksheet worksheetLocal, int rowNumber)
        {
            List<string> newRow = new List<string>();

            foreach (var cellAddres in cellsAdresses)
            {
                var cell = worksheet.Cells[$"{cellAddres}{rowNumber}"];
                var cellLocal = worksheetLocal.Cells[$"{cellAddres}{rowNumber}"];

                newRow.Add(cell.Text);

                cellLocal.Value = cell.Value;
            }

            return newRow;
        }
    }
}