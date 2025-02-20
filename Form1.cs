using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Открыть диалог для выбора файла Word
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents|*.doc;*.docx";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Вызываем метод конвертации
                ConvertWordToExcel(openFileDialog.FileName);
            }
        }
        //Конвертация из экселя в ворд
        private void ConvertWordToExcel(string wordFilePath)
        {
            // Создаем объекты для Word и Excel
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelWorkbook = null;
            Worksheet excelWorksheet = null;

            try
            {
                // Открываем документ Word
                Document wordDoc = wordApp.Documents.Open(wordFilePath);
                wordApp.Visible = false;

                // Создаем новый файл Excel
                excelWorkbook = excelApp.Workbooks.Add();
                excelWorksheet = (Worksheet)excelWorkbook.Worksheets[1];

                // Получаем таблицы из документа Word
                Tables tables = wordDoc.Tables;

                if (tables.Count > 0)
                {
                    Table wordTable = tables[1]; // Берем первую таблицу
                    int rowCount = wordTable.Rows.Count; // Общее количество строк
                    int colCount = wordTable.Columns.Count; // Общее количество столбцов

                    // Копируем заголовки в первую строку Excel
                    for (int col = 1; col <= colCount; col++)
                    {
                        string headerText = wordTable.Cell(1, col).Range.Text;
                        headerText = headerText.Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();
                        excelWorksheet.Cells[1, col] = headerText;
                    }

                    int excelRow = 2; // Начнем с 2-й строки для Excel
                                      // Копируем данные из таблицы Word в Excel
                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Проверяем, если во втором столбце этой строки цифра 2
                        if (wordTable.Cell(row, 2).Range.Text.Trim('\r', '\a') == "2")
                        {
                            continue; // Пропускаем эту строку
                        }

                        for (int col = 1; col <= colCount; col++)
                        {
                            string cellText = wordTable.Cell(row, col).Range.Text.Trim('\r', '\a');
                            excelWorksheet.Cells[excelRow, col] = cellText;
                        }

                        // Добавляем порядковый номер в первый столбец Excel
                        excelWorksheet.Cells[excelRow, 1] = excelRow - 1;
                        excelRow++; // Увеличиваем счетчик строк для Excel
                    }

                    // Сохраняем файл Excel
                    string excelFilePath = System.IO.Path.ChangeExtension(wordFilePath, ".xlsx");
                    excelWorkbook.SaveAs(excelFilePath);

                    MessageBox.Show($"Конвертация завершена. Файл сохранен: {excelFilePath}");
                    UpdatePublicationCounts(excelFilePath); //по видам изданий за все время
                    UpdatePublicationCountsByYear(excelFilePath); //по видам изданий по годам
                    UpdateSpecializationCounts(excelFilePath); //по специальности за все время
                    UpdateSpecializationCountsByYear(excelFilePath); //по специальности по годам
                    UpdateKVCounts(excelFilePath); //по кварталу за все время
                    UpdateуKVCountsByYear(excelFilePath); //по кварталу по годам
                    UpdateAuthorPublicationCounts(excelFilePath); //по авторам
                    ConvertExcelToCsv(excelFilePath);
                }
                else
                {
                    MessageBox.Show("В документе не найдено таблиц.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
            finally
            {
                // Закрываем все приложения
                if (excelWorkbook != null)
                {
                    excelWorkbook.Close(false); // Закрываем книгу без сохранения
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
                if (wordApp != null)
                {
                    wordApp.Quit();
                }
            }
        }
        //конвертация из экселя в csv
        private void ConvertExcelToCsv(string excelFilePath)
        {
            // Создаем экземпляр приложения Excel
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false; // Открываем Excel в фоновом режиме

            try
            {
                // Открываем существующий Excel файл
                Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);

                // Получаем первый лист в книге
                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];

                // Определяем путь для сохранения CSV файла
                string csvFilePath = System.IO.Path.ChangeExtension(excelFilePath, ".csv");

                // Сохраняем лист в формате CSV
                worksheet.SaveAs(csvFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);

                // Закрываем книгу без сохранения
                workbook.Close(false);
                MessageBox.Show($"Файл успешно сохранен в формате CSV: {csvFilePath}");

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
            finally
            {
                // Обязательно закрываем приложение Excel
                excelApp.Quit();
            }

        }

        // Словарь для хранения количества каждого вида издания
        private Dictionary<string, int> publicationCounts = new Dictionary<string, int>();
        
        private void UpdatePublicationCounts(string excelFilePath)
        {
            // Создаём новое приложение Excel, которое не отображается
            var excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false // Не показывать Excel интерфейс
            };

            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                workbook = excelApp.Workbooks.Open(excelFilePath);
                worksheet = workbook.Sheets[1]; // Получаем первый лист (можно указать также по имени)

                int row = 2; // Начинаем со второй строки (предположим, что первая строка - заголовки)

                while (worksheet.Cells[row, 4].Value != null) // Пока есть данные в четвертой колонке
                {
                    string publicationType = worksheet.Cells[row, 4].Value.ToString().Trim().ToLower();

                    if (!string.IsNullOrEmpty(publicationType))
                    {
                        if (publicationCounts.ContainsKey(publicationType))
                        {
                            publicationCounts[publicationType]++;
                        }
                        else
                        {
                            publicationCounts[publicationType] = 1;
                        }
                    }

                    row++;
                }
            }
            finally
            {
                // Освобождаем ресурсы
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            // Обновляем диаграмму
            UpdateChart();
        }

        private void UpdateChart()
        {
            // Очищаем старые данные с диаграммы
            chart1.Series[0].Points.Clear();
            chart1.Series[0].LegendText = "Вид издания";

            // Добавляем новые данные
            foreach (var kvp in publicationCounts)
            {
                chart1.Series[0].Points.AddXY(kvp.Key, kvp.Value);
            }

            // Устанавливаем равные интервалы по оси Y
            chart1.ChartAreas[0].AxisY.Interval = 1; // Интервал по оси Y равен 1
            chart1.ChartAreas[0].AxisY.IsStartedFromZero = true; // Начинать ось Y с 0

            // Обновляем диаграмму
            chart1.Invalidate();
        }

        // Словарь для хранения количества каждого вида издания по годам
        private Dictionary<string, Dictionary<string, int>> yearlyPublicationCounts = new Dictionary<string, Dictionary<string, int>>();

        private void UpdatePublicationCountsByYear(string excelFilePath)
        {
            // Создаём новое приложение Excel, которое не отображается
            var excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false // Не показывать Excel интерфейс
            };

            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                workbook = excelApp.Workbooks.Open(excelFilePath);
                worksheet = workbook.Sheets[1]; // Получаем первый лист (можно указать также по имени)

                int row = 2; // Начинаем с второй строки (предположим, что первая строка - заголовки)

                while (worksheet.Cells[row, 4].Value != null) // Пока есть данные в четвертой колонке (тип издания)
                {
                    string publicationType = worksheet.Cells[row, 4].Value.ToString().Trim().ToLower();

                    // Получаем год из имени файла
                    var yearMatch = Regex.Match(Path.GetFileName(excelFilePath), @"(\d{4})");
                    if (!yearMatch.Success)
                    {
                        row++;
                        continue; // Если год не найден, пропускаем эту строку
                    }

                    string year = yearMatch.Value;

                    // Инициализация словаря для данного года, если он не существует
                    if (!yearlyPublicationCounts.ContainsKey(year))
                    {
                        yearlyPublicationCounts[year] = new Dictionary<string, int>();
                    }

                    // Увеличиваем счетчик для данного вида публикации и года
                    if (!string.IsNullOrEmpty(publicationType))
                    {
                        if (yearlyPublicationCounts[year].ContainsKey(publicationType))
                        {
                            yearlyPublicationCounts[year][publicationType]++;
                        }
                        else
                        {
                            yearlyPublicationCounts[year][publicationType] = 1;
                        }
                    }

                    row++;
                }
            }
            finally
            {
                // Освобождаем ресурсы
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            // Обновляем диаграмму
            UpdateChart2();
        }

        private void UpdateChart2()
        {
            // Очищаем старые данные с диаграммы
            chart2.Series.Clear();

            // Добавляем данные в series
            foreach (var yearData in yearlyPublicationCounts)
            {
                string year = yearData.Key;
                foreach (var kvp in yearData.Value)
                {
                    var series = new System.Windows.Forms.DataVisualization.Charting.Series
                    {
                        Name = $"{year} - {kvp.Key}", // Имя серии будет год + вид издания
                        IsValueShownAsLabel = true,
                        ChartType = SeriesChartType.Column // Или другой тип диаграммы
                    };

                    double xValue = Convert.ToDouble(year); // Значение по оси X - год
                    series.Points.AddXY(year, kvp.Value); // Добавляем точку на диаграмму
                    chart2.Series.Add(series); // Добавляем серию на диаграмму
                }
            }

            // Устанавливаем равные интервалы по оси Y
            chart2.ChartAreas[0].AxisY.Interval = 1; // Интервал по оси Y равен 1
            chart2.ChartAreas[0].AxisY.IsStartedFromZero = true; // Начинать ось Y с 0

            // Скрываем значения по оси X
            chart2.ChartAreas[0].AxisX.LabelStyle.Enabled = false; // Отключаем отображение меток на оси X

            // Обновляем диаграмму
            chart2.Invalidate();
        }

        // Словарь для хранения количества каждой специальности
        private Dictionary<string, int> specializationCounts = new Dictionary<string, int>();

        private void UpdateSpecializationCounts(string excelFilePath)
        {
            // Создаём новое приложение Excel, которое не отображается
            var excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false // Не показывать Excel интерфейс
            };

            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                workbook = excelApp.Workbooks.Open(excelFilePath);
                worksheet = workbook.Sheets[1]; // Получаем первый лист (можно указать также по имени)

                int row = 2; // Начинаем со второй строки (предположим, что первая строка - заголовки)

                while (worksheet.Cells[row, 5].Value != null) // Пока есть данные в 5 колонке
                {
                    string publicationType = worksheet.Cells[row, 5].Value.ToString().Trim().ToLower();

                    if (!string.IsNullOrEmpty(publicationType))
                    {
                        if (specializationCounts.ContainsKey(publicationType))
                        {
                            specializationCounts[publicationType]++;
                        }
                        else
                        {
                            specializationCounts[publicationType] = 1;
                        }
                    }

                    row++;
                }
            }
            finally
            {
                // Освобождаем ресурсы
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            // Обновляем диаграмму
            UpdateChart3();
        }

        private void UpdateChart3()
        {
            // Очищаем старые данные с диаграммы
            chart3.Series[0].Points.Clear();
            chart3.Series[0].LegendText = "Специальность";

            // Добавляем новые данные
            foreach (var kvp in specializationCounts)
            {
                chart3.Series[0].Points.AddXY(kvp.Key, kvp.Value);
            }

            // Устанавливаем равные интервалы по оси Y
            chart3.ChartAreas[0].AxisY.Interval = 1; // Интервал по оси Y равен 1
            chart3.ChartAreas[0].AxisY.IsStartedFromZero = true; // Начинать ось Y с 0

            // Обновляем диаграмму
            chart3.Invalidate();
        }

        // Словарь для хранения количества каждой специальности по годам
        private Dictionary<string, Dictionary<string, int>> yearlyPSpezializationCounts = new Dictionary<string, Dictionary<string, int>>();

        private void UpdateSpecializationCountsByYear(string excelFilePath)
        {
            // Создаём новое приложение Excel, которое не отображается
            var excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false // Не показывать Excel интерфейс
            };

            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                workbook = excelApp.Workbooks.Open(excelFilePath);
                worksheet = workbook.Sheets[1]; // Получаем первый лист (можно указать также по имени)

                int row = 2; // Начинаем с второй строки (предположим, что первая строка - заголовки)

                while (worksheet.Cells[row, 5].Value != null) // Пока есть данные в 5 колонке (специальность)
                {
                    string specType = worksheet.Cells[row, 5].Value.ToString().Trim().ToLower();

                    // Получаем год из имени файла
                    var yearMatch = Regex.Match(Path.GetFileName(excelFilePath), @"(\d{4})");
                    if (!yearMatch.Success)
                    {
                        row++;
                        continue; // Если год не найден, пропускаем эту строку
                    }

                    string year = yearMatch.Value;

                    // Инициализация словаря для данного года, если он не существует
                    if (!yearlyPSpezializationCounts.ContainsKey(year))
                    {
                        yearlyPSpezializationCounts[year] = new Dictionary<string, int>();
                    }

                    // Увеличиваем счетчик для данной специальности и года
                    if (!string.IsNullOrEmpty(specType))
                    {
                        if (yearlyPSpezializationCounts[year].ContainsKey(specType))
                        {
                            yearlyPSpezializationCounts[year][specType]++;
                        }
                        else
                        {
                            yearlyPSpezializationCounts[year][specType] = 1;
                        }
                    }

                    row++;
                }
            }
            finally
            {
                // Освобождаем ресурсы
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            // Обновляем диаграмму
            UpdateChart4();
        }

        private void UpdateChart4()
        {
            // Очищаем старые данные с диаграммы
            chart4.Series.Clear();

            // Добавляем данные в series
            foreach (var yearData in yearlyPSpezializationCounts)
            {
                string year = yearData.Key;
                foreach (var kvp in yearData.Value)
                {
                    var series = new System.Windows.Forms.DataVisualization.Charting.Series
                    {
                        Name = $"{year} - {kvp.Key}", // Имя серии будет год + специальность
                        IsValueShownAsLabel = true,
                        ChartType = SeriesChartType.Column // Или другой тип диаграммы
                    };

                    double xValue = Convert.ToDouble(year); // Значение по оси X - год
                    series.Points.AddXY(year, kvp.Value); // Добавляем точку на диаграмму
                    chart4.Series.Add(series); // Добавляем серию на диаграмму
                }
            }

            // Устанавливаем равные интервалы по оси Y
            chart4.ChartAreas[0].AxisY.Interval = 1; // Интервал по оси Y равен 1
            chart4.ChartAreas[0].AxisY.IsStartedFromZero = true; // Начинать ось Y с 0

            // Скрываем значения по оси X
            chart4.ChartAreas[0].AxisX.LabelStyle.Enabled = false; // Отключаем отображение меток на оси X

            // Обновляем диаграмму
            chart4.Invalidate();
        }
        // Словарь для хранения количества каждого квартала
        private Dictionary<string, int> kvCounts = new Dictionary<string, int>();

        private void UpdateKVCounts(string excelFilePath)
        {
            // Создаём новое приложение Excel, которое не отображается
            var excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false // Не показывать Excel интерфейс
            };

            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                workbook = excelApp.Workbooks.Open(excelFilePath);
                worksheet = workbook.Sheets[1]; // Получаем первый лист (можно указать также по имени)

                int row = 2; // Начинаем со второй строки (предположим, что первая строка - заголовки)
                             // Получаем общее количество колонок в листе
                int columnCount = worksheet.UsedRange.Columns.Count;

                // Выбираем номер строки в зависимости от количества колонок
                int targetColumn = (columnCount == 9) ? 9 : (columnCount == 8) ? 8 : -1;

                while (worksheet.Cells[row, targetColumn].Value != null) // Пока есть данные в 8 колонке
                {
                    string publicationType = worksheet.Cells[row, targetColumn].Value.ToString().Trim().ToLower();

                    if (!string.IsNullOrEmpty(publicationType))
                    {
                        if (kvCounts.ContainsKey(publicationType))
                        {
                            kvCounts[publicationType]++;
                        }
                        else
                        {
                            kvCounts[publicationType] = 1;
                        }
                    }

                    row++;
                }
            }
            finally
            {
                // Освобождаем ресурсы
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            // Обновляем диаграмму
            UpdateChart5();
        }

        private void UpdateChart5()
        {
            // Очищаем старые данные с диаграммы
            chart5.Series[0].Points.Clear();
            chart5.Series[0].LegendText = "Квартал сдачи";

            // Добавляем новые данные
            foreach (var kvp in kvCounts)
            {
                chart5.Series[0].Points.AddXY(kvp.Key, kvp.Value);
            }

            // Устанавливаем равные интервалы по оси Y
            chart5.ChartAreas[0].AxisY.Interval = 1; // Интервал по оси Y равен 1
            chart5.ChartAreas[0].AxisY.IsStartedFromZero = true; // Начинать ось Y с 0

            // Обновляем диаграмму
            chart5.Invalidate();
        }

        // Словарь для хранения количества каждой специальности по годам
        private Dictionary<string, Dictionary<string, int>> yearlyKVCounts = new Dictionary<string, Dictionary<string, int>>();

        private void UpdateуKVCountsByYear(string excelFilePath)
        {
            // Создаём новое приложение Excel, которое не отображается
            var excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false // Не показывать Excel интерфейс
            };

            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                workbook = excelApp.Workbooks.Open(excelFilePath);
                worksheet = workbook.Sheets[1]; // Получаем первый лист (можно указать также по имени)

                int row = 2; // Начинаем с второй строки (предположим, что первая строка - заголовки)

                int columnCount = worksheet.UsedRange.Columns.Count;

                // Выбираем номер строки в зависимости от количества колонок
                int targetColumn = (columnCount == 9) ? 9 : (columnCount == 8) ? 8 : -1;

                while (worksheet.Cells[row, targetColumn].Value != null) // Пока есть данные в 8 колонке (квартал сдачи)
                {
                    string specType = worksheet.Cells[row, targetColumn].Value.ToString().Trim().ToLower();

                    // Получаем год из имени файла
                    var yearMatch = Regex.Match(Path.GetFileName(excelFilePath), @"(\d{4})");
                    if (!yearMatch.Success)
                    {
                        row++;
                        continue; // Если год не найден, пропускаем эту строку
                    }

                    string year = yearMatch.Value;

                    // Инициализация словаря для данного года, если он не существует
                    if (!yearlyKVCounts.ContainsKey(year))
                    {
                        yearlyKVCounts[year] = new Dictionary<string, int>();
                    }

                    // Увеличиваем счетчик для данной специальности и года
                    if (!string.IsNullOrEmpty(specType))
                    {
                        if (yearlyKVCounts[year].ContainsKey(specType))
                        {
                            yearlyKVCounts[year][specType]++;
                        }
                        else
                        {
                            yearlyKVCounts[year][specType] = 1;
                        }
                    }

                    row++;
                }
            }
            finally
            {
                // Освобождаем ресурсы
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            // Обновляем диаграмму
            UpdateChart6();
        }

        private void UpdateChart6()
        {
            // Очищаем старые данные с диаграммы
            chart6.Series.Clear();

            // Добавляем данные в series
            foreach (var yearData in yearlyKVCounts)
            {
                string year = yearData.Key;
                foreach (var kvp in yearData.Value)
                {
                    var series = new System.Windows.Forms.DataVisualization.Charting.Series
                    {
                        Name = $"{year} - {kvp.Key} квартал", // Имя серии будет год + квартал
                        IsValueShownAsLabel = true,
                        ChartType = SeriesChartType.Column // Или другой тип диаграммы
                    };

                    double xValue = Convert.ToDouble(year); // Значение по оси X - год
                    series.Points.AddXY(year, kvp.Value); // Добавляем точку на диаграмму
                    chart6.Series.Add(series); // Добавляем серию на диаграмму
                }
            }

            // Устанавливаем равные интервалы по оси Y
            chart6.ChartAreas[0].AxisY.Interval = 1; // Интервал по оси Y равен 1
            chart6.ChartAreas[0].AxisY.IsStartedFromZero = true; // Начинать ось Y с 0

            // Скрываем значения по оси X
            chart6.ChartAreas[0].AxisX.LabelStyle.Enabled = false; // Отключаем отображение меток на оси X

            // Обновляем диаграмму
            chart6.Invalidate();
        }

        // Словарь для хранения сумм по каждому автору
        private Dictionary<string, int> authorPublicationCounts = new Dictionary<string, int>();

        private void UpdateAuthorPublicationCounts(string excelFilePath)
        {
            // Создаём новое приложение Excel, которое не отображается
            var excelApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = false // Не показывать Excel интерфейс
            };

            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                workbook = excelApp.Workbooks.Open(excelFilePath);
                worksheet = workbook.Sheets[1]; // Получаем первый лист (можно указать также по имени)

                int row = 2; // Начинаем со второй строки (предположим, что первая строка - заголовки)

                int columnCount = worksheet.UsedRange.Columns.Count;

                // Выбираем номер строки в зависимости от количества колонок
                int targetColumn = (columnCount == 9) ? 7 : (columnCount == 8) ? 6 : -1;

                while (worksheet.Cells[row, targetColumn].Value != null) // Пока есть данные во втором столбце (Авторы)
                {
                    string author = worksheet.Cells[row, 2].Value.ToString().Trim().ToLower(); // Берём значение автора
                    int volume = (int)(worksheet.Cells[row, targetColumn].Value ?? 0); // Берём объем (6-й столбец)

                    if (!string.IsNullOrEmpty(author))
                    {
                        if (authorPublicationCounts.ContainsKey(author))
                        {
                            authorPublicationCounts[author] += volume; // Суммируем объём
                        }
                        else
                        {
                            authorPublicationCounts[author] = volume; // Инициализируем суммой объема
                        }
                    }

                    row++;
                }
            }
            finally
            {
                // Освобождаем ресурсы
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            // Обновляем диаграмму
            UpdateAuthorChart();
        }

        private void UpdateAuthorChart()
        {
            // Очищаем старые данные с диаграммы
            chart7.Series[0].Points.Clear();
            chart7.Series[0].LegendText = "Авторы";

            // Добавляем новые данные
            foreach (var kvp in authorPublicationCounts)
            {
                chart7.Series[0].Points.AddXY(kvp.Key, kvp.Value);
            }

            // Устанавливаем равные интервалы по оси Y
            chart7.ChartAreas[0].AxisY.Interval = 1; // Интервал по оси Y равен 1
            chart7.ChartAreas[0].AxisY.IsStartedFromZero = true; // Начинать ось Y с 0

            // Установка оси X для вертикальных меток
            chart7.ChartAreas[0].AxisX.IsLabelAutoFit = false; // Отключаем автоматическую подгонку
            chart7.ChartAreas[0].AxisX.LabelStyle.Angle = -90; // Устанавливаем угол наклона меток в -90 градусов
            chart7.ChartAreas[0].AxisX.LabelStyle.Enabled = true; // Убедитесь, что метки включены

            // Установка явных меток для всех точек
            chart7.ChartAreas[0].AxisX.Interval = 1; // Указываем, что каждая метка должна отображаться
            chart7.ChartAreas[0].AxisX.MajorTickMark.Interval = 1; // Убедитесь, что главные деления устанавливаются для каждой метки

            // Настройка легенды
            chart7.Legends[0].Docking = Docking.Bottom;

            // Обновляем диаграмму
            chart7.Invalidate();
        }
    }

}

