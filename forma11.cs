using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace WindowsFormsApp1
{
    public partial class Printed_form_diagrams : Form
    {
        public Printed_form_diagrams()
        {
            InitializeComponent();
        }

        //словарь для хранения пути файла
        private Dictionary<string, bool> processedFiles = new Dictionary<string, bool>();
        private void export_to_csv_xslx_button_Click(object sender, EventArgs e)
        {
            // Открыть диалог для выбора файлов Word
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents|*.doc;*.docx";
            openFileDialog.Multiselect = true; //  Разрешаем выбор нескольких файлов

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Получаем массив выбранных файлов
                string[] selectedFiles = openFileDialog.FileNames;

                // Перебираем выбранные файлы и вызываем метод конвертации для каждого
                foreach (string filePath in selectedFiles)
                {
                    ConvertWordToCsv(filePath); // Преобразуем файл в CSV
                }
            }
        }

        //Конвертация из экселя в ворд
        private void ConvertWordToCsv(string wordFilePath)
        {
            // Создаем объекты для Word
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            try
            {
                // Открываем документ Word
                Document wordDoc = wordApp.Documents.Open(wordFilePath);
                wordApp.Visible = false;

                // Получаем таблицы из документа Word
                Tables tables = wordDoc.Tables;

                if (tables.Count > 0)
                {
                    Table wordTable = tables[1]; // Берем первую таблицу
                    int rowCount = wordTable.Rows.Count; // Общее количество строк
                    int colCount = wordTable.Columns.Count; // Общее количество столбцов

                    // Создаем строку для заголовков CSV
                    string headerRow = "";
                    for (int col = 1; col <= colCount; col++)
                    {
                        string headerText = wordTable.Cell(1, col).Range.Text;
                        headerText = headerText.Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();

                        // Экранируем запятые в заголовках, если необходимо
                        if (headerText.Contains(","))
                        {
                            headerText = "\"" + headerText + "\"";
                        }

                        headerRow += headerText;
                        if (col < colCount)
                        {
                            headerRow += ",";
                        }
                    }

                    // Создаем список для строк CSV данных
                    List<string> csvRows = new List<string>();
                    csvRows.Add(headerRow); // Добавляем строку заголовков

                    // Преобразуем данные из таблицы Word в строки CSV
                    int counter = 1; // Инициализируем счетчик

                    for (int row = 2; row <= rowCount; row++)
                    {

                        string dataRow = counter.ToString() + ","; // Добавляем счетчик в начало строки CSV и запятую
                        counter++; // Увеличиваем счетчик для следующей строки

                        for (int col = 2; col <= colCount; col++) // Начинаем со второго столбца Word
                        {
                            string cellText = wordTable.Cell(row, col).Range.Text.Replace("\r", "").Replace("\n", "").Replace("\a", "").Trim();

                            // Экранируем запятые в данных, если необходимо
                            if (cellText.Contains(","))
                            {
                                cellText = "\"" + cellText + "\"";
                            }

                            dataRow += cellText;
                            if (col < colCount)
                            {
                                dataRow += ",";
                            }
                        }

                        csvRows.Add(dataRow);
                    }

                    // Определяем путь для сохранения CSV файла
                    string csvFilePath = System.IO.Path.ChangeExtension(wordFilePath, ".csv");

                    // Записываем данные в CSV файл, указав кодировку UTF-8
                    File.WriteAllLines(csvFilePath, csvRows, Encoding.UTF8);

                    // Проверяем, был ли файл уже обработан
                    if (!processedFiles.ContainsKey(csvFilePath))
                    {
                        // Добавляем путь к файлу в словарь, чтобы отметить его как обработанный
                        processedFiles.Add(csvFilePath, true);
                    }

                    //MessageBox.Show($"CSV файл успешно создан: {csvFilePath}");
                    UpdatePublicationCounts(csvFilePath); //по видам изданий за все время
                    UpdatePublicationCountsByYear(csvFilePath); //по видам изданий по годам
                    UpdateSpecializationCounts(csvFilePath); //по специальности за все время
                    UpdateSpecializationCountsByYear(csvFilePath); //по специальности по годам
                    UpdateKVCounts(csvFilePath); //по кварталу за все время
                    UpdateуKVCountsByYear(csvFilePath); //по кварталу по годам
                    UpdateAuthorPublicationCounts(csvFilePath); //по авторам

                    // Закрываем документ Word
                    wordDoc.Close();
                    wordApp.Quit();
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
                // Гарантированно освобождаем ресурсы Word
                if (wordApp != null)
                {
                    wordApp.Quit();
                }
            }
        }

        // Словарь для хранения количества каждого вида издания
        private Dictionary<string, int> publicationCounts = new Dictionary<string, int>();

        //парсин строк
        private static List<string[]> ParseCsvLine(string line)
        {
            // Инициализируем список для хранения результатов
            List<string[]> result = new List<string[]>();

            // Флаг для отслеживания, находимся ли мы внутри кавычек
            bool inQuotes = false;

            // Список для хранения значений текущей строки
            List<string> currentValues = new List<string>();

            // Строка для хранения текущего значения между разделителями
            string currentValue = "";

            // Проходим по каждому символу в строке
            for (int i = 0; i < line.Length; i++)
            {
                // Если встречаем кавычку, меняем состояние флага inQuotes
                if (line[i] == '"')
                {
                    inQuotes = !inQuotes; // Меняем статус: если были внутри кавычек - выходим, если нет - заходим
                }
                // Если встретили запятую и не находимся внутри кавычек
                else if (line[i] == ',' && !inQuotes)
                {
                    // Добавляем текущее значение в список значений
                    currentValues.Add(currentValue);
                    // Сбрасываем текущее значение для следующего
                    currentValue = "";
                }
                else
                {
                    // Если это не запятая, добавляем символ к текущему значению
                    currentValue += line[i];
                }
            }
            // Добавляем последнее значение после завершения цикла
            currentValues.Add(currentValue);

            // Добавляем массив значений в результат
            result.Add(currentValues.ToArray());

            return result; // Возвращаем список массивов значений
        }

        private void UpdatePublicationCounts(string csvFilePath)
        {
            try
            {
                // Считываем все строки из CSV файла
                var lines = File.ReadAllLines(csvFilePath);

                // Пропускаем первую строку (заголовки) и получаем остальные строки данных
                var dataLines = lines.Skip(1);

                int rowNumber = 2; // Начинаем со второй строки (индекс 1 после Skip(1))

                // Проходим по каждой строке данных
                foreach (var line in dataLines)
                {
                    // Проверяем, является ли текущая строка второй строкой и содержит ли она "1,2,3"
                    if (rowNumber == 2 && line.Contains("1,2,3"))
                    {
                        rowNumber++; //Переходим к 3-й строке
                        continue; // Пропускаем текущую итерацию цикла
                    }

                    // Парсим строку с помощью метода ParseCsvLine
                    List<string[]> parsedLine = ParseCsvLine(line);

                    // Проверяем, что парсинг дал результат и в первой строке есть как минимум 4 элемента
                    if (parsedLine.Any() && parsedLine.First().Length >= 4)
                    {
                        // Получаем тип публикации из четвертого элемента (индекс 3) и обрабатываем его
                        string publicationType = parsedLine.First()[3].Trim().ToLower();

                        // Проверяем, не является ли строка пустой
                        if (!string.IsNullOrEmpty(publicationType))
                        {
                            // Если тип публикации уже есть в словаре, увеличиваем его счетчик
                            if (publicationCounts.ContainsKey(publicationType))
                            {
                                publicationCounts[publicationType]++;
                            }
                            // Если его нет, добавляем новый тип с начальным счетчиком
                            else
                            {
                                publicationCounts[publicationType] = 1;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Обработка ошибок чтения файла
               MessageBox.Show($"Ошибка при чтении CSV файла: {ex.Message}");
            }

            // Обновляем график или пользовательский интерфейс, если это необходимо
            UpdatePublicationCountsChart();
        }

        private void UpdatePublicationCountsChart()
        {
            // Очищаем старые данные с диаграммы
            num_publication_by_type_all_time_chart.Series[0].Points.Clear();
            num_publication_by_type_all_time_chart.Series[0].LegendText = "Вид издания";

            // Добавляем новые данные
            foreach (var kvp in publicationCounts)
            {
                num_publication_by_type_all_time_chart.Series[0].Points.AddXY(kvp.Key, kvp.Value);
            }

            // Устанавливаем равные интервалы по оси Y
            num_publication_by_type_all_time_chart.ChartAreas[0].AxisY.Interval = 1; // Интервал по оси Y равен 1
            num_publication_by_type_all_time_chart.ChartAreas[0].AxisY.IsStartedFromZero = true; // Начинать ось Y с 0

            // Обновляем диаграмму
            num_publication_by_type_all_time_chart.Invalidate();
        }

        // Словарь для хранения количества каждого вида издания по годам
        private Dictionary<string, Dictionary<string, int>> yearlyPublicationCounts = new Dictionary<string, Dictionary<string, int>>();

        private void UpdatePublicationCountsByYear(string csvFilePath)
        {
            try
            {
                // Получаем год из имени файла
                var yearMatch = Regex.Match(Path.GetFileName(csvFilePath), @"(\d{4})");
                if (!yearMatch.Success)
                {
                    MessageBox.Show($"Предупреждение: Год не найден в имени файла '{csvFilePath}'.");
                    return; // Если год не найден, прекращаем обработку файла
                }

                string year = yearMatch.Value;

                // Читаем все строки из CSV файла
                var lines = File.ReadAllLines(csvFilePath);

                // Пропускаем первую строку (заголовки), если она есть
                var dataLines = lines.Skip(1);

                int rowNumber = 2; // Начинаем со второй строки (индекс 1 после Skip(1))

                // Проходим по каждой строке данных
                foreach (var line in dataLines)
                {
                    // Проверяем, является ли текущая строка второй строкой и содержит ли она "1,2,3"
                    if (rowNumber == 2 && line.Contains("1,2,3"))
                    {
                        rowNumber++; //Переходим к 3-й строке
                        continue; // Пропускаем текущую итерацию цикла
                    }
                    // Разбираем строку CSV с учетом экранирования кавычками
                    List<string[]> parsedLine = ParseCsvLine(line);

                    // Проверяем, что строка была успешно разобрана и содержит как минимум 4 столбца
                    if (parsedLine.Any() && parsedLine.First().Length >= 4)
                    {
                        string publicationType = parsedLine.First()[3].Trim().ToLower();

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
                    }
                    else
                    {
                        MessageBox.Show($"Предупреждение: Строка '{line}' содержит менее 4 столбцов или не была корректно разобрана.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении CSV файла: {ex.Message}");
            }

            // Обновляем диаграмму
            UpdatePublicationCountsByYearChart();
        }

        private void UpdatePublicationCountsByYearChart()
        {
            // Очищаем старые данные с диаграммы
            num_publication_by_type_year_chart.Series.Clear();

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
                    num_publication_by_type_year_chart.Series.Add(series); // Добавляем серию на диаграмму
                }
            }

            // Устанавливаем равные интервалы по оси Y
            num_publication_by_type_year_chart.ChartAreas[0].AxisY.Interval = 1; // Интервал по оси Y равен 1
            num_publication_by_type_year_chart.ChartAreas[0].AxisY.IsStartedFromZero = true; // Начинать ось Y с 0

            // Скрываем значения по оси X
            num_publication_by_type_year_chart.ChartAreas[0].AxisX.LabelStyle.Enabled = false; // Отключаем отображение меток на оси X

            // Обновляем диаграмму
            num_publication_by_type_year_chart.Invalidate();
        }

        // Словарь для хранения количества каждой специальности
        private Dictionary<string, int> specializationCounts = new Dictionary<string, int>();

        private void UpdateSpecializationCounts(string csvFilePath)
        {
            try
            {
                // Читаем все строки из CSV файла
                var lines = File.ReadAllLines(csvFilePath);

                // Пропускаем первую строку (заголовки), если она есть
                var dataLines = lines.Skip(1);

                int rowNumber = 2; // Начинаем со второй строки (индекс 1 после Skip(1))

                // Проходим по каждой строке данных
                foreach (var line in dataLines)
                {
                    // Проверяем, является ли текущая строка второй строкой и содержит ли она "1,2,3"
                    if (rowNumber == 2 && line.Contains("1,2,3"))
                    {
                        rowNumber++; //Переходим к 3-й строке
                        continue; // Пропускаем текущую итерацию цикла
                    }
                    // Разбираем строку CSV с учетом экранирования кавычками
                    List<string[]> parsedLine = ParseCsvLine(line);

                    // Проверяем, что строка была успешно разобрана и содержит как минимум 6 столбцов (индекс 5)
                    if (parsedLine.Any() && parsedLine.First().Length >= 6)
                    {
                        string specializationType = parsedLine.First()[4].Trim().ToLower(); // Получаем значение из 5-го столбца

                        // Проверяем, что строка не пустая
                        if (!string.IsNullOrEmpty(specializationType))
                        {
                            // Обновляем словарь specializationCounts
                            if (specializationCounts.ContainsKey(specializationType))
                            {
                                specializationCounts[specializationType]++;
                            }
                            else
                            {
                                specializationCounts[specializationType] = 1;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show($"Предупреждение: Строка '{line}' содержит менее 6 столбцов или не была корректно разобрана.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении CSV файла: {ex.Message}");
            }

            // Обновляем диаграмму
            UpdateSpecializationCountsChart();
        }

        private void UpdateSpecializationCountsChart()
        {
            // Очищаем старые данные с диаграммы
            num_publication_by_specialization_all_time_chart.Series[0].Points.Clear();
            num_publication_by_specialization_all_time_chart.Series[0].LegendText = "Специальность";

            // Добавляем новые данные
            foreach (var kvp in specializationCounts)
            {
                num_publication_by_specialization_all_time_chart.Series[0].Points.AddXY(kvp.Key, kvp.Value);
            }

            // Устанавливаем равные интервалы по оси Y
            num_publication_by_specialization_all_time_chart.ChartAreas[0].AxisY.Interval = 1; // Интервал по оси Y равен 1
            num_publication_by_specialization_all_time_chart.ChartAreas[0].AxisY.IsStartedFromZero = true; // Начинать ось Y с 0

            // Обновляем диаграмму
            num_publication_by_specialization_all_time_chart.Invalidate();
        }

        // Словарь для хранения количества каждой специальности по годам
        private Dictionary<string, Dictionary<string, int>> yearlyPSpezializationCounts = new Dictionary<string, Dictionary<string, int>>();

        private void UpdateSpecializationCountsByYear(string csvFilePath)
        {
            try
            {
                // Получаем год из имени файла
                var yearMatch = Regex.Match(Path.GetFileName(csvFilePath), @"(\d{4})");
                if (!yearMatch.Success)
                {
                    MessageBox.Show($"Предупреждение: Год не найден в имени файла '{csvFilePath}'.");
                    return; // Если год не найден, прекращаем обработку файла
                }

                string year = yearMatch.Value;

                // Читаем все строки из CSV файла
                var lines = File.ReadAllLines(csvFilePath);

                // Пропускаем первую строку (заголовки), если она есть
                var dataLines = lines.Skip(1);

                int rowNumber = 2; // Начинаем со второй строки (индекс 1 после Skip(1))

                // Проходим по каждой строке данных
                foreach (var line in dataLines)
                {
                    // Проверяем, является ли текущая строка второй строкой и содержит ли она "1,2,3"
                    if (rowNumber == 2 && line.Contains("1,2,3"))
                    {
                        rowNumber++; //Переходим к 3-й строке
                        continue; // Пропускаем текущую итерацию цикла
                    }
                    // Разбираем строку CSV с учетом экранирования кавычками
                    List<string[]> parsedLine = ParseCsvLine(line);

                    // Проверяем, что строка была успешно разобрана и содержит как минимум 6 столбцов (индекс 5)
                    if (parsedLine.Any() && parsedLine.First().Length >= 6)
                    {
                        string specType = parsedLine.First()[4].Trim().ToLower(); // Получаем значение из 5-го столбца

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
                    }
                    else
                    {
                        MessageBox.Show($"Предупреждение: Строка '{line}' содержит менее 6 столбцов или не была корректно разобрана.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении CSV файла: {ex.Message}");
            }

            // Обновляем диаграмму
            UpdateSpecializationCountsByYearChart();
        }

        private void UpdateSpecializationCountsByYearChart()
        {
            // Очищаем старые данные с диаграммы
            num_publication_by_specialization_year_chart.Series.Clear();

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
                    num_publication_by_specialization_year_chart.Series.Add(series); // Добавляем серию на диаграмму
                }
            }

            // Устанавливаем равные интервалы по оси Y
            num_publication_by_specialization_year_chart.ChartAreas[0].AxisY.Interval = 1; // Интервал по оси Y равен 1
            num_publication_by_specialization_year_chart.ChartAreas[0].AxisY.IsStartedFromZero = true; // Начинать ось Y с 0

            // Скрываем значения по оси X
            num_publication_by_specialization_year_chart.ChartAreas[0].AxisX.LabelStyle.Enabled = false; // Отключаем отображение меток на оси X

            // Обновляем диаграмму
            num_publication_by_specialization_year_chart.Invalidate();
        }
        // Словарь для хранения количества каждого квартала
        private Dictionary<string, int> kvCounts = new Dictionary<string, int>();

        private void UpdateKVCounts(string csvFilePath)
        {
            try
            {
                // Читаем все строки из CSV файла
                var lines = File.ReadAllLines(csvFilePath);

                // Пропускаем первую строку (заголовки), если она есть
                var dataLines = lines.Skip(1);

                // Определяем, сколько колонок ожидаем (7 или 8), исходя из данных
                int expectedColumnCount = (lines.FirstOrDefault()?.Split(',').Length == 8) ? 8 :
                                          (lines.FirstOrDefault()?.Split(',').Length == 9) ? 9 : -1;

                if (expectedColumnCount == -1)
                {
                    MessageBox.Show($"Предупреждение: Не удалось определить количество столбцов (ожидается 8 или 9) в файле '{csvFilePath}'.");
                    return;
                }

                int rowNumber = 2; // Начинаем со второй строки (индекс 1 после Skip(1))

                // Проходим по каждой строке данных
                foreach (var line in dataLines)
                {
                    // Проверяем, является ли текущая строка второй строкой и содержит ли она "1,2,3"
                    if (rowNumber == 2 && line.Contains("1,2,3"))
                    {
                        rowNumber++; //Переходим к 3-й строке
                        continue; // Пропускаем текущую итерацию цикла
                    }
                    // Разбираем строку CSV с учетом экранирования кавычками
                    List<string[]> parsedLine = ParseCsvLine(line);

                    // Проверяем, что строка была успешно разобрана и содержит ожидаемое количество столбцов
                    if (parsedLine.Any() && parsedLine.First().Length == expectedColumnCount)
                    {
                        int targetColumnIndex = expectedColumnCount - 1; // Индекс столбца (от 0)
                        string publicationType = parsedLine.First()[targetColumnIndex].Trim().ToLower(); // Получаем значение из целевого столбца

                        // Проверяем, что строка не пустая
                        if (!string.IsNullOrEmpty(publicationType))
                        {
                            // Обновляем словарь kvCounts
                            if (kvCounts.ContainsKey(publicationType))
                            {
                                kvCounts[publicationType]++;
                            }
                            else
                            {
                                kvCounts[publicationType] = 1;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show($"Предупреждение: Строка '{line}' содержит неверное количество столбцов или не была корректно разобрана.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении CSV файла: {ex.Message}");
            }

            // Обновляем диаграмму
            UpdateKVCountsChart();
        }

        private void UpdateKVCountsChart()
        {
            // Очищаем старые данные с диаграммы
            num_publication_by_quarter_all_time_chart.Series[0].Points.Clear();
            num_publication_by_quarter_all_time_chart.Series[0].LegendText = "Квартал сдачи";

            // Добавляем новые данные
            foreach (var kvp in kvCounts)
            {
                num_publication_by_quarter_all_time_chart.Series[0].Points.AddXY(kvp.Key, kvp.Value);
            }

            // Устанавливаем равные интервалы по оси Y
            num_publication_by_quarter_all_time_chart.ChartAreas[0].AxisY.Interval = 1; // Интервал по оси Y равен 1
            num_publication_by_quarter_all_time_chart.ChartAreas[0].AxisY.IsStartedFromZero = true; // Начинать ось Y с 0

            // Обновляем диаграмму
            num_publication_by_quarter_all_time_chart.Invalidate();
        }

        // Словарь для хранения количества каждой специальности по годам
        private Dictionary<string, Dictionary<string, int>> yearlyKVCounts = new Dictionary<string, Dictionary<string, int>>();

        private void UpdateуKVCountsByYear(string csvFilePath)
        {
            try
            {
                // Получаем год из имени файла
                var yearMatch = Regex.Match(Path.GetFileName(csvFilePath), @"(\d{4})");
                if (!yearMatch.Success)
                {
                    MessageBox.Show($"Предупреждение: Год не найден в имени файла '{csvFilePath}'.");
                    return; // Если год не найден, прекращаем обработку файла
                }

                string year = yearMatch.Value;

                // Читаем все строки из CSV файла
                var lines = File.ReadAllLines(csvFilePath);

                // Пропускаем первую строку (заголовки), если она есть
                var dataLines = lines.Skip(1);

                // Определяем, сколько колонок ожидаем (8 или 9), исходя из данных
                int expectedColumnCount = (lines.FirstOrDefault()?.Split(',').Length == 8) ? 8 :
                                          (lines.FirstOrDefault()?.Split(',').Length == 9) ? 9 : -1;

                if (expectedColumnCount == -1)
                {
                    MessageBox.Show($"Предупреждение: Не удалось определить количество столбцов (ожидается 8 или 9) в файле '{csvFilePath}'.");
                    return;
                }

                int rowNumber = 2; // Начинаем со второй строки (индекс 1 после Skip(1))

                // Проходим по каждой строке данных
                foreach (var line in dataLines)
                {
                    // Проверяем, является ли текущая строка второй строкой и содержит ли она "1,2,3"
                    if (rowNumber == 2 && line.Contains("1,2,3"))
                    {
                        rowNumber++; //Переходим к 3-й строке
                        continue; // Пропускаем текущую итерацию цикла
                    }
                    // Разбираем строку CSV с учетом экранирования кавычками
                    List<string[]> parsedLine = ParseCsvLine(line);

                    // Проверяем, что строка была успешно разобрана и содержит ожидаемое количество столбцов
                    if (parsedLine.Any() && parsedLine.First().Length == expectedColumnCount)
                    {
                        int targetColumnIndex = expectedColumnCount - 1; // Индекс целевого столбца (от 0)
                        string specType = parsedLine.First()[targetColumnIndex].Trim().ToLower(); // Получаем значение из целевого столбца

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
                    }
                    else
                    {
                        MessageBox.Show($"Предупреждение: Строка '{line}' содержит неверное количество столбцов или не была корректно разобрана.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении CSV файла: {ex.Message}");
            }

            // Обновляем диаграмму
            UpdateKVCountsByYearChart();
        }

        private void UpdateKVCountsByYearChart()
        {
            // Очищаем старые данные с диаграммы
            num_publication_by_quarter_year_chart.Series.Clear();

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
                    num_publication_by_quarter_year_chart.Series.Add(series); // Добавляем серию на диаграмму
                }
            }

            // Устанавливаем равные интервалы по оси Y
            num_publication_by_quarter_year_chart.ChartAreas[0].AxisY.Interval = 1; // Интервал по оси Y равен 1
            num_publication_by_quarter_year_chart.ChartAreas[0].AxisY.IsStartedFromZero = true; // Начинать ось Y с 0

            // Скрываем значения по оси X
            num_publication_by_quarter_year_chart.ChartAreas[0].AxisX.LabelStyle.Enabled = false; // Отключаем отображение меток на оси X

            // Обновляем диаграмму
            num_publication_by_quarter_year_chart.Invalidate();
        }

        // Словарь для хранения сумм по каждому автору
        private Dictionary<string, int> authorPublicationCounts = new Dictionary<string, int>();

        private void UpdateAuthorPublicationCounts(string csvFilePath)
        {
            try
            {
                // Читаем все строки из CSV файла
                var lines = File.ReadAllLines(csvFilePath);

                // Пропускаем первую строку (заголовки), если она есть
                var dataLines = lines.Skip(1);

                // Определяем, сколько колонок ожидаем (8 или 9), исходя из данных
                int expectedColumnCount = (lines.FirstOrDefault()?.Split(',').Length == 8) ? 8 :
                                          (lines.FirstOrDefault()?.Split(',').Length == 9) ? 9 : -1;

                if (expectedColumnCount == -1)
                {
                    MessageBox.Show($"Предупреждение: Не удалось определить количество столбцов (ожидается 8 или 9) в файле '{csvFilePath}'.");
                    return;
                }
                int rowNumber = 2; // Начинаем со второй строки (индекс 1 после Skip(1))

                // Проходим по каждой строке данных
                foreach (var line in dataLines)
                {
                    // Проверяем, является ли текущая строка второй строкой и содержит ли она "1,2,3"
                    if (rowNumber == 2 && line.Contains("1,2,3"))
                    {
                        rowNumber++; //Переходим к 3-й строке
                        continue; // Пропускаем текущую итерацию цикла
                    }
                    // Разбираем строку CSV с учетом экранирования кавычками
                    List<string[]> parsedLine = ParseCsvLine(line);

                    // Проверяем, что строка была успешно разобрана и содержит ожидаемое количество столбцов
                    if (parsedLine.Any() && parsedLine.First().Length == expectedColumnCount)
                    {
                        int authorColumnIndex = 1; // Второй столбец (индекс 1) - Автор
                        int volumeColumnIndex = (expectedColumnCount == 8) ? 5 : (expectedColumnCount == 9) ? 6 : -1;

                        string author = parsedLine.First()[authorColumnIndex].Trim().ToLower();
                        string volumeString = parsedLine.First()[volumeColumnIndex]; // Объем как строка

                        if (!string.IsNullOrEmpty(author) && !string.IsNullOrEmpty(volumeString) && int.TryParse(volumeString, out int volume))
                        {
                            if (authorPublicationCounts.ContainsKey(author))
                            {
                                authorPublicationCounts[author] += volume; // Суммируем объем
                            }
                            else
                            {
                                authorPublicationCounts[author] = volume; // Инициализируем суммой объема
                            }
                        }
                        else
                        {
                            MessageBox.Show($"Предупреждение: Не удалось прочитать данные из строки '{line}'.");
                        }
                    }
                    else
                    {
                        MessageBox.Show($"Предупреждение: Строка '{line}' содержит неверное количество столбцов или не была корректно разобрана.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении CSV файла: {ex.Message}");
            }

            // Обновляем диаграмму
            UpdateAuthorChart();
        }

        private void UpdateAuthorChart()
        {
            // Очищаем старые данные с диаграммы
            num_publication_by_volume_author_page_all_time_chart.Series[0].Points.Clear();
            num_publication_by_volume_author_page_all_time_chart.Series[0].LegendText = "Авторы";

            // Добавляем новые данные
            foreach (var kvp in authorPublicationCounts)
            {
                num_publication_by_volume_author_page_all_time_chart.Series[0].Points.AddXY(kvp.Key, kvp.Value);
            }

            // Устанавливаем равные интервалы по оси Y
            num_publication_by_volume_author_page_all_time_chart.ChartAreas[0].AxisY.Interval = 1; // Интервал по оси Y равен 1
            num_publication_by_volume_author_page_all_time_chart.ChartAreas[0].AxisY.IsStartedFromZero = true; // Начинать ось Y с 0

            // Установка оси X для вертикальных меток
            num_publication_by_volume_author_page_all_time_chart.ChartAreas[0].AxisX.IsLabelAutoFit = false; // Отключаем автоматическую подгонку
            num_publication_by_volume_author_page_all_time_chart.ChartAreas[0].AxisX.LabelStyle.Angle = -90; // Устанавливаем угол наклона меток в -90 градусов
            num_publication_by_volume_author_page_all_time_chart.ChartAreas[0].AxisX.LabelStyle.Enabled = true; // Убедитесь, что метки включены

            // Установка явных меток для всех точек
            num_publication_by_volume_author_page_all_time_chart.ChartAreas[0].AxisX.Interval = 1; // Указываем, что каждая метка должна отображаться

            // Настройка легенды
            num_publication_by_volume_author_page_all_time_chart.Legends[0].Docking = Docking.Bottom;

            // Обновляем диаграмму
            num_publication_by_volume_author_page_all_time_chart.Invalidate();
        }

        private void ConvertCsvToExcel(string csvFilePath, string excelFilePath)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Add();
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1]; // Получаем первый лист

                int rowNumber = 1;
                foreach (string line in File.ReadLines(csvFilePath))
                {
                    List<string[]> values = ParseCsvLine(line); // Парсим строку CSV

                    int colNumber = 1;
                    foreach (string value in values[0])
                    {
                        worksheet.Cells[rowNumber, colNumber] = value; // Записываем значение в ячейку Excel
                        colNumber++;
                    }

                    rowNumber++;
                }

                workbook.SaveAs(excelFilePath); // Сохраняем Excel-файл
                excelApp.Quit(); // Закрываем Excel
            }
            finally
            {
                // Освобождаем COM-объекты для предотвращения утечек памяти
                if (worksheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (workbook != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                if (excelApp != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                worksheet = null;
                workbook = null;
                excelApp = null;

                GC.Collect(); // Вызываем сборщик мусора для немедленной очистки памяти
                GC.WaitForPendingFinalizers();
            }
        }

        public void ConvertCsvFilesToExcel()
        {
            // Создаем копию ключей словаря, чтобы избежать проблем при изменении словаря в цикле.
            List<string> filePaths = new List<string>(processedFiles.Keys);
            try
            {
                foreach (string filePath in filePaths)
                {
                    string excelFilePath = Path.ChangeExtension(filePath, ".xlsx"); // Создаем имя для Excel-файла

                    // Проверяем, существует ли Excel-файл
                    if (File.Exists(excelFilePath))
                    {
                        MessageBox.Show($"Файл {excelFilePath} уже существует.");
                        continue; // Переходим к следующему файлу
                    }

                    ConvertCsvToExcel(filePath, excelFilePath); // Вызываем метод конвертации
                }

                MessageBox.Show($"Файлы успешно конвертированы.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при конвертации: {ex.Message}");
            }
        }

        private void export_to_excel_button_Click(object sender, EventArgs e)
        {
            ConvertCsvFilesToExcel();
        }

        public void ConvertCsvFilesToWordTables()
        {
            // Создаем копию ключей словаря, чтобы избежать проблем при изменении словаря в цикле.
            List<string> filePaths = new List<string>(processedFiles.Keys);
            try
            {
                foreach (string filePath in filePaths)
                {

                    // Создаем имя для Word-файла
                    string wordFilePath = Path.Combine(
                        Path.GetDirectoryName(filePath),
                        Path.GetFileNameWithoutExtension(filePath) + " из csv.docx"
                    );

                    // Проверяем, существует ли Word-файл
                    if (File.Exists(wordFilePath))
                    {
                        MessageBox.Show($"Файл {wordFilePath} уже существует. ");
                        continue; // Переходим к следующему файлу
                    }

                    ConvertCsvToWordTable(filePath, wordFilePath); // Вызываем метод конвертации
                    MessageBox.Show($"Файлы успешно конвертированы.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при конвертации: {ex.Message}");

            }
        }

        private void ConvertCsvToWordTable(string csvFilePath, string wordFilePath)
        {
            Microsoft.Office.Interop.Word.Application wordApp = null;
            Microsoft.Office.Interop.Word.Document wordDoc = null;

            try
            {
                wordApp = new Microsoft.Office.Interop.Word.Application();
                wordDoc = wordApp.Documents.Add();

                List<string[]> data = new List<string[]>();

                // Читаем данные из CSV-файла и парсим их
                foreach (string line in File.ReadLines(csvFilePath))
                {
                    List<string[]> values = ParseCsvLine(line);
                    data.Add(values[0]); // values[0], т.к. ParseCsvLine возвращает List<string[]> с одним элементом
                }

                // Создаем таблицу в Word
                if (data.Count > 0)
                {
                    Microsoft.Office.Interop.Word.Table table = wordDoc.Tables.Add(wordApp.Selection.Range, data.Count, data[0].Length, Missing.Value, Missing.Value);

                    // Заполняем таблицу данными
                    for (int row = 0; row < data.Count; row++)
                    {
                        for (int col = 0; col < data[row].Length; col++)
                        {
                            table.Cell(row + 1, col + 1).Range.Text = data[row][col];
                        }
                    }

                    // Добавляем границы ко всей таблице
                    table.Borders.Enable = 1; // Включаем отображение границ
                    table.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle; // Стиль внешней границы
                    table.Borders.OutsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth050pt; // Толщина внешней границы
                    table.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle; // Стиль внутренней границы
                    table.Borders.InsideLineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth050pt; // Толщина внутренней границы

                    wordDoc.SaveAs2(wordFilePath);
                }
                else
                {
                    MessageBox.Show($"CSV файл {csvFilePath} пуст. Таблица не создана.");
                }
            }
            finally
            {
                // Закрываем и освобождаем объекты
                if (wordDoc != null)
                {
                    wordDoc.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc);
                    wordDoc = null;
                }
                if (wordApp != null)
                {
                    wordApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                    wordApp = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void export_to_word_button_Click(object sender, EventArgs e)
        {
            ConvertCsvFilesToWordTables();
        }
    }

}

