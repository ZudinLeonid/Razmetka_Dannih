using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using System.Windows.Input;
using System.ComponentModel;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;



namespace WpfApp3
{
    public partial class MainWindow : Window
    {
        private ReadExcel excel;
        private ScottPlot.Plottable.ScatterPlot MyScatterPlot;
        private List<ScottPlot.Plottable.MarkerPlot> HighlightedPoint;
        private ScottPlot.Drawing.Font font;

        private void WpfPlot1_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            (double mouseCoordX, double mouseCoordY) = WpfPlot1.GetMouseCoordinates();

            if (HighlightedPoint.Count < 4)
            {
                HighlightedPoint.Add(WpfPlot1.Plot.AddPoint(Math.Round(mouseCoordX, 3), 0));

                HighlightedPoint[HighlightedPoint.Count - 1].MarkerShape = ScottPlot.MarkerShape.openCircle;
                HighlightedPoint[HighlightedPoint.Count - 1].Color = System.Drawing.Color.Red;
                HighlightedPoint[HighlightedPoint.Count - 1].MarkerLineWidth = 7;
                HighlightedPoint[HighlightedPoint.Count - 1].MarkerSize = 7;
                HighlightedPoint[HighlightedPoint.Count - 1].Text = $"{HighlightedPoint.Count}";
                HighlightedPoint[HighlightedPoint.Count - 1].TextFont = font;
                HighlightedPoint[HighlightedPoint.Count - 1].IsVisible = true;
            }

            if (Point1.Text == "")
                Point1.Text = HighlightedPoint[HighlightedPoint.Count - 1].X.ToString().Replace(',', '.');
            else if (Point2.Text == "")
                Point2.Text = HighlightedPoint[HighlightedPoint.Count - 1].X.ToString().Replace(',', '.');
            else if (Point3.Text == "")
                Point3.Text = HighlightedPoint[HighlightedPoint.Count - 1].X.ToString().Replace(',', '.');
            else if (Point4.Text == "")
                Point4.Text = HighlightedPoint[HighlightedPoint.Count - 1].X.ToString().Replace(',', '.');


            WpfPlot1.Render();
        }

        private void NextDetector_Click(object sender, RoutedEventArgs e)
        {
            if (!(excel.CurrentDetector <= excel.detector[excel.CurrentRow][excel.CurrentDetector].Length))
                return;

            excel.CurrentDetector++;
            CurrentDetectorTextBox.Text = $"{ excel.CurrentDetector}";
            excel.DrowCell(excel.CurrentRow, excel.CurrentDetector);
            MyScatterPlot.Update(excel.time, excel.amplitude);
            WpfPlot1.Plot.AxisAuto();


            string currentCellValue = excel.userTime[excel.CurrentRow][excel.CurrentDetector];
            if (currentCellValue != "")
            {
                currentCellValue = currentCellValue.Replace(",", "|");
                currentCellValue = currentCellValue.Replace(".", ",");
                double[] values = Array.ConvertAll(currentCellValue.Split('|'), double.Parse);

                for (int i = 0; i < values.Length; i++)
                {
                    HighlightedPoint.Add(WpfPlot1.Plot.AddPoint(values[i], 0));

                    HighlightedPoint[i].MarkerShape = ScottPlot.MarkerShape.openCircle;
                    HighlightedPoint[i].Color = System.Drawing.Color.Red;
                    HighlightedPoint[i].MarkerLineWidth = 7;
                    HighlightedPoint[i].MarkerSize = 7;
                    HighlightedPoint[i].Text = $"{i + 1}";
                    HighlightedPoint[i].TextFont = font;
                    HighlightedPoint[i].IsVisible = true;
                }

                for (int i = 0; i < HighlightedPoint.Count; i++)
                    ((TextBox)FindName($"Point{i + 1}")).Text = HighlightedPoint[i]?.X.ToString();
            }
            else
            {
                HighlightedPoint.ForEach(point => WpfPlot1.Plot.Remove(point));
                HighlightedPoint.Clear();

                for (int i = 0; i < 4; i++) { ((TextBox)FindName($"Point{i + 1}")).Text = ""; }
            }

            WpfPlot1.Refresh();
        }
        private void PrevDetector_Click(object sender, RoutedEventArgs e)
        {
            if (!(excel.CurrentDetector - 1 >= 0))
                return;

            excel.CurrentDetector--;
            CurrentDetectorTextBox.Text = $"{ excel.CurrentDetector}";
            excel.DrowCell(excel.CurrentRow, excel.CurrentDetector);
            MyScatterPlot.Update(excel.time, excel.amplitude);
            WpfPlot1.Plot.AxisAuto();

            string currentCellValue = excel.userTime[excel.CurrentRow][excel.CurrentDetector];
            if (currentCellValue != "")
            {
                currentCellValue = currentCellValue.Replace(",", "|");
                currentCellValue = currentCellValue.Replace(".", ",");
                double[] values = Array.ConvertAll(currentCellValue.Split('|'), double.Parse);

                HighlightedPoint.ForEach(point => WpfPlot1.Plot.Remove(point));
                HighlightedPoint.Clear();
                HighlightedPoint = new List<ScottPlot.Plottable.MarkerPlot>();

                for (int i = 0; i < values.Length; i++)
                {
                    HighlightedPoint.Add(WpfPlot1.Plot.AddPoint(values[i], 0));

                    HighlightedPoint[i].MarkerShape = ScottPlot.MarkerShape.openCircle;
                    HighlightedPoint[i].Color = System.Drawing.Color.Red;
                    HighlightedPoint[i].MarkerLineWidth = 7;
                    HighlightedPoint[i].MarkerSize = 7;
                    HighlightedPoint[i].Text = $"{i + 1}";
                    HighlightedPoint[i].TextFont = font;
                    HighlightedPoint[i].IsVisible = true;
                }

                for (int i = 0; i < HighlightedPoint.Count; i++)
                    ((TextBox)FindName($"Point{i + 1}")).Text = HighlightedPoint[i]?.X.ToString();
            }
            else
            {
                HighlightedPoint.ForEach(point => WpfPlot1.Plot.Remove(point));
                HighlightedPoint.Clear();

                for (int i = 0; i < 4; i++) { ((TextBox)FindName($"Point{i + 1}")).Text = ""; }
            }

            WpfPlot1.Refresh();
        }

        private void NextRow_Click(object sender, RoutedEventArgs e)
        {
            if (!(excel.CurrentRow + 1 < excel.rowsCount))
                return;

            excel.CurrentRow++;
            CurrentRowTextBox.Text = $"{excel.CurrentRow}";
            excel.DrowCell(excel.CurrentRow, excel.CurrentDetector);
            MyScatterPlot.Update(excel.time, excel.amplitude);
            WpfPlot1.Plot.AxisAuto();

            string currentCellValue = excel.userTime[excel.CurrentRow][excel.CurrentDetector];
            if (currentCellValue != "")
            {
                currentCellValue = currentCellValue.Replace(",", "|");
                currentCellValue = currentCellValue.Replace(".", ",");
                double[] values = Array.ConvertAll(currentCellValue.Split('|'), double.Parse);

                HighlightedPoint.ForEach(point => WpfPlot1.Plot.Remove(point));
                HighlightedPoint.Clear();
                HighlightedPoint = new List<ScottPlot.Plottable.MarkerPlot>();

                for (int i = 0; i < values.Length; i++)
                {
                    HighlightedPoint.Add(WpfPlot1.Plot.AddPoint(values[i], 0));

                    HighlightedPoint[i].MarkerShape = ScottPlot.MarkerShape.openCircle;
                    HighlightedPoint[i].Color = System.Drawing.Color.Red;
                    HighlightedPoint[i].MarkerLineWidth = 7;
                    HighlightedPoint[i].MarkerSize = 7;
                    HighlightedPoint[i].Text = $"{i + 1}";
                    HighlightedPoint[i].TextFont = font;
                    HighlightedPoint[i].IsVisible = true;
                }

                for (int i = 0; i < HighlightedPoint.Count; i++)
                    ((TextBox)FindName($"Point{i + 1}")).Text = HighlightedPoint[i]?.X.ToString();
            }
            else
            {
                HighlightedPoint.ForEach(point => WpfPlot1.Plot.Remove(point));
                HighlightedPoint.Clear();

                for (int i = 0; i < 4; i++) { ((TextBox)FindName($"Point{i + 1}")).Text = ""; }
            }

            WpfPlot1.Refresh();
        }
        private void PrevRow_Click(object sender, RoutedEventArgs e)
        {
            if (!(excel.CurrentRow - 1 >= 0))
                return;

            excel.CurrentRow--;
            CurrentRowTextBox.Text = $"{excel.CurrentRow}";
            excel.DrowCell(excel.CurrentRow, excel.CurrentDetector);
            MyScatterPlot.Update(excel.time, excel.amplitude);
            WpfPlot1.Plot.AxisAuto();

            string currentCellValue = excel.userTime[excel.CurrentRow][excel.CurrentDetector];
            if (currentCellValue != "")
            {
                currentCellValue = currentCellValue.Replace(",", "|");
                currentCellValue = currentCellValue.Replace(".", ",");
                double[] values = Array.ConvertAll(currentCellValue.Split('|'), double.Parse);

                HighlightedPoint.ForEach(point => WpfPlot1.Plot.Remove(point));
                HighlightedPoint.Clear();
                HighlightedPoint = new List<ScottPlot.Plottable.MarkerPlot>();

                for (int i = 0; i < values.Length; i++)
                {
                    HighlightedPoint.Add(WpfPlot1.Plot.AddPoint(values[i], 0));

                    HighlightedPoint[i].MarkerShape = ScottPlot.MarkerShape.openCircle;
                    HighlightedPoint[i].Color = System.Drawing.Color.Red;
                    HighlightedPoint[i].MarkerLineWidth = 7;
                    HighlightedPoint[i].MarkerSize = 7;
                    HighlightedPoint[i].Text = $"{i + 1}";
                    HighlightedPoint[i].TextFont = font;
                    HighlightedPoint[i].IsVisible = true;
                }

                for (int i = 0; i < HighlightedPoint.Count; i++)
                    ((TextBox)FindName($"Point{i + 1}")).Text = HighlightedPoint[i]?.X.ToString();
            }
            else
            {
                HighlightedPoint.ForEach(point => WpfPlot1.Plot.Remove(point));
                HighlightedPoint.Clear();

                for (int i = 0; i < 4; i++) { ((TextBox)FindName($"Point{i + 1}")).Text = ""; }
            }

            WpfPlot1.Refresh();
        }

        private void CurrentRowTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
                return;

            if (!int.TryParse(CurrentRowTextBox.Text, out int newRow))
                return;

            if (newRow < 0 || newRow >= 115)
                return;

            excel.CurrentRow = newRow;
            excel.DrowCell(excel.CurrentRow, excel.CurrentDetector);
            MyScatterPlot.Update(excel.time, excel.amplitude);
            WpfPlot1.Plot.AxisAuto();
            WpfPlot1.Refresh();
        }
        private void CurrentDetectorTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
                return;

            if (!int.TryParse(CurrentDetectorTextBox.Text, out int newDetector))
                return;

            if (newDetector < 0 || newDetector >= 400)
                return;

            excel.CurrentDetector = newDetector;
            excel.DrowCell(excel.CurrentRow, excel.CurrentDetector);
            MyScatterPlot.Update(excel.time, excel.amplitude);
            WpfPlot1.Plot.AxisAuto();
            WpfPlot1.Refresh();
        }


        private void DeleteAllButton_Click(object sender, RoutedEventArgs e)
        {
            HighlightedPoint.ForEach(point => WpfPlot1.Plot.Remove(point));
            HighlightedPoint.Clear();

            Point1.Text = "";
            Point2.Text = "";
            Point3.Text = "";
            Point4.Text = "";

            WpfPlot1.Refresh();
        }
        private void SavePointsButton_Click(object sender, RoutedEventArgs e)
        {
            string output = $"{Point1.Text},{Point2.Text},{Point3.Text},{Point4.Text}";
            if (output.EndsWith(",")) { output = output.TrimEnd(','); }
            excel.userTime[excel.CurrentRow][excel.CurrentDetector] = output;
        }

        private void SaveInFileButton_Click(object sender, RoutedEventArgs e) { excel.SaveFile(); }

        public MainWindow()
        {
            InitializeComponent();

            HighlightedPoint = new List<ScottPlot.Plottable.MarkerPlot>();

            excel = new ReadExcel();
            excel.OpenFile();

            font = new ScottPlot.Drawing.Font();
            font.Size = 30;
            font.Alignment = ScottPlot.Alignment.LowerLeft;
            font.Bold = true;
            font.Color = System.Drawing.Color.Black;

            excel.DrowCell(excel.CurrentRow, excel.CurrentDetector);

            CurrentRowTextBox.Text = $"{excel.CurrentRow}";
            CurrentDetectorTextBox.Text = $"{ excel.CurrentDetector}";

            DataContext = excel;
            CurrentRowTextBox.KeyUp += CurrentRowTextBox_KeyUp;
            CurrentDetectorTextBox.KeyUp += CurrentDetectorTextBox_KeyUp;

            MyScatterPlot = WpfPlot1.Plot.AddScatter(excel.time, excel.amplitude);
            WpfPlot1.Plot.Title("Установите 4 точки");
            WpfPlot1.Plot.XLabel("Время");
            WpfPlot1.Plot.YLabel("Амплитуда");

            WpfPlot1.Plot.AddHorizontalLine(0);

            WpfPlot1.MouseLeftButtonUp += WpfPlot1_MouseLeftButtonUp;

            WpfPlot1.Refresh();
        }
    }

    public class ReadExcel : INotifyPropertyChanged
    {
        public double[] time;
        public double[] amplitude;
        public string[][] detector;
        public List<List<string>> userTime;
        public int rowsCount;
        public int columnsCount;
        private int _currentRow;
        private int _currentDetector;
        
        public int CurrentRow
        {
            get => _currentRow;
            set
            {
                if (_currentRow != value)
                {
                    _currentRow = value;
                    OnPropertyChanged(nameof(CurrentRow));
                }
            }
        }
        public int CurrentDetector
        {
            get => _currentDetector;
            set
            {
                if (_currentDetector != value)
                {
                    _currentDetector = value;
                    OnPropertyChanged(nameof(CurrentRow));
                }
            }
        }

        public ReadExcel()
        {
            _currentRow = 0;
            _currentDetector = 0;
            rowsCount = -1;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void OpenFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV Files (*.csv)|*.csv";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;

                CheckNumsOfRows(filePath);
                ReadFromCSV(filePath);
            }
        }
        public void SaveFile()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "CSV Files (*.csv)|*.csv";
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;

                WriteDataToCSV(filePath);
            }
        }
        public void WriteDataToCSV(string filePath)
        {
            using (StreamWriter writer = new StreamWriter(filePath))
            {
                StringBuilder sbColumnNames = new StringBuilder("row;position;");
                
                for (int i = 0; i < 400; i++)
                {
                    string detectorNum = "detector_" + i.ToString() + ";";
                    sbColumnNames.Append(detectorNum);
                }

                sbColumnNames.Remove(sbColumnNames.Length - 1, 1);
                writer.WriteLine(sbColumnNames.ToString());

                StringBuilder dataRow = new StringBuilder();
                int position = 130675;
                for (int row = 0; row < rowsCount; row++)
                {
                    position += row;
                    dataRow.Append($"{row};{position};");
                    for (int column = 0; column < 400; column++)
                    {
                        if (userTime[row][column] != "")
                            dataRow.Append($"{userTime[row][column]};");
                        else
                            dataRow.Append("--;");
                    }
                    dataRow.Remove(dataRow.Length - 1, 1);
                    writer.WriteLine(dataRow.ToString());
                    dataRow.Clear();
                }
            }
        }

        public void CheckNumsOfRows(string filePath)
        {
            using (StreamReader reader = new StreamReader(filePath))
            {
                while (!reader.EndOfStream)
                {
                    columnsCount = (new string[reader.ReadLine().Split(';').Length - 2]).Length;
                    rowsCount++;
                }
            }
        }

        public void DrowCell(int rowNum, int detectorNum)
        {
            if (detector[rowNum][detectorNum] == "--") { detectorNum++; }

            string cell = detector[rowNum][detectorNum];

            string[] elementPairs = cell.Split('|');

            string[] strTime = new string[elementPairs.Length];
            string[] strAmplitude = new string[elementPairs.Length];

            int index = 0;
            for (int i = 0; i < elementPairs.Length; i++)
            {
                string[] pair = elementPairs[i].Split(':');

                if (pair[0] == "--") { index--; break; }

                strTime[i + index] = pair[0];
                strAmplitude[i + index] = pair[1];
            }

            string[] filteredTime = strTime.Where(element => element != null).ToArray();
            time = filteredTime.Select(element => double.Parse(element)).ToArray();

            string[] filteredAmplitude = strAmplitude.Where(element => element != null).ToArray();
            amplitude = filteredAmplitude.Select(element => double.Parse(element)).ToArray();
        }

        public void ReadFromCSV(string filePath)
        {
            using (StreamReader reader = new StreamReader(filePath))
            {
                if (!reader.EndOfStream) { reader.ReadLine(); }

                detector = new string[rowsCount][];

                int index = 0;

                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    line = line.Replace(",", "|");
                    line = line.Replace(".", ",");

                    detector[index] = new string[line.Split(';').Length - 2];
                    Array.Copy(line.Split(';'), 2, detector[index], 0, line.Split(';').Length - 2);

                    index++;
                }

                userTime = new List<List<string>>();
                for (int i = 0; i < rowsCount; i++)
                {
                    List<string> row = new List<string>();
                    for (int j = 0; j < detector[0].Length; j++)
                    {
                        row.Add("");
                    }
                    userTime.Add(row);
                }
            }
        }

        public void ReadFromMultiplyCells(string excelPath, string firstRangeElement, string secondRangeElement)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelPath);
            Excel.Worksheet worksheet = workbook.Sheets["Sheet1"];
            Excel.Range range = worksheet.Range[firstRangeElement + ":" + secondRangeElement];

            object[,] values = (object[,])range.Value;

            int rowCount = values.GetLength(0);
            int colCount = values.GetLength(1);

            object value;
            string cell_value = "";
            for (int row = 1; row <= rowCount; row++) {
                for (int col = 1; col <= colCount; col++) {
                    value = values[row, col];
                    cell_value += value.ToString();
                }
            }

            string multiple_spaces = @"\s+";
            cell_value = string.Join("", cell_value.Split('[', ']', '\n'));
            cell_value = cell_value.Replace(".", ",");
            cell_value = Regex.Replace(cell_value, multiple_spaces, " ");
            cell_value = cell_value.Trim();

            string[] substrings = cell_value.Split(' ');

            bool flag = false;
            time = new double[substrings.Length / 2 + 1];
            amplitude = new double[substrings.Length / 2 + 1];
            int j = -32;
            for (int i = 0; i < substrings.Length; i++) {
                if (i % 32 == 0) {
                    flag = !flag;
                    j += 32;
                }

                if (flag == true) {
                    time[i - j] = double.Parse(substrings[i]);
                }
                if (flag == false) {
                    amplitude[i - j] = double.Parse(substrings[i]);
                }
            }
        }

        public void ReadFromSingleCell(string excelPath, int sheetIndex, string cell)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelPath);
            Excel.Worksheet worksheet = workbook.Sheets[sheetIndex];
            Excel.Range range = worksheet.Range[cell];

            string cell_value = range.Value.ToString();

            // Удаляем символы '[', ']', '\n'
            cell_value = string.Join("", cell_value.Split('[', ']', '\n'));

            // Форматируем строку (25.8 -> 25,8), чтобы потом можно было преобразовать string в double
            cell_value = cell_value.Replace(".", ",");

            // Избавляемся от многократных пробелов между значениями
            string multiple_spaces = @"\s+";
            cell_value = Regex.Replace(cell_value, multiple_spaces, " ");
            cell_value = cell_value.Trim();

            // Разбиваем строку, на подстроки. В каждой подстроке хранится одно значение
            string[] substrings = cell_value.Split(' ');

            bool flag = false;
            time = new double[substrings.Length / 2 + 1];
            amplitude = new double[substrings.Length / 2 + 1];
            int j = -32;
            for (int i = 0; i < substrings.Length; i++) {
                if (i % 32 == 0) {
                    flag = !flag;
                    j += 32;
                }

                if (flag == true) {
                    time[i - j] = double.Parse(substrings[i]);
                }
                if (flag == false) {
                    amplitude[i - j] = double.Parse(substrings[i]);
                }
            }
        }

        public void OpenExcelFile(string excelPath, string firstRangeElement, string secondRangeElement)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelPath);
            Excel.Worksheet worksheet = workbook.Sheets["Sheet1"];
            Excel.Range range = worksheet.Range[firstRangeElement + ":" + secondRangeElement];

            string cell_value = range.Value.ToString();

            cell_value = cell_value.Replace("\n", " ");

            char[] trimChars = { '[', ']' };
            cell_value = cell_value.Trim(trimChars);
            cell_value = cell_value.Trim();

            string multiple_spaces = @"\s+";
            cell_value = Regex.Replace(cell_value, multiple_spaces, " ");
            cell_value = cell_value.Replace(".", ",");

            string[] substrings = cell_value.Split(' ');
        }
    }

}