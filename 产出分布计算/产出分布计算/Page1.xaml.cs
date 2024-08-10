using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Drawing;
using System.Windows.Input;
using System.Diagnostics;
using System;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using static 产出分布计算.Page1;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using System.Collections.Concurrent;
using System.IO.MemoryMappedFiles;
using System.ComponentModel;
using System.Windows.Shapes;
using System.Data;
using Newtonsoft.Json;
using System.Net;
using System.Text.Json;
using System.Diagnostics;

namespace 产出分布计算
{
    /// <summary>
    /// Page1.xaml 的交互逻辑
    /// </summary>
    public partial class Page1 : Page
    {
        private int _progress;
        public int Progress
        {
            get { return _progress; }
            set
            {
                _progress = value;
                OnPropertyChanged("Progress");
            }
        }
        public class TextBoxContent
        {
            public string BinName { get; set; }
            public string BinSetting { get; set; }
            public string Dimension { get; set; }
            public string FilePath { get; set; }
            public string filenameSuffix { get; set; }
            public string Para1 { get; set; }
            public string Para1Min { get; set; }
            public string Para1Rta { get; set; }
            public string Para1Num { get; set; }
            public string Fix1num { get; set; }
            public string Para2 { get; set; }
            public string Para2Min { get; set; }
            public string Para2Rta { get; set; }
            public string Para2Num { get; set; }
            public string Fix2num { get; set; }
            public string Para3 { get; set; }
            public string Para3Min { get; set; }
            public string Para3Rta { get; set; }
            public string Para3Num { get; set; }
            public string Fix3num { get; set; }
            public string Para4 { get; set; }
            public string Para4Min { get; set; }
            public string Para4Rta { get; set; }
            public string Para4Num { get; set; }
            public string Fix4num { get; set; }
        }

        private readonly object fileWriteLock = new object(); // 用于保护文件写入操作的锁对象
        private readonly object parameterlockObject = new object();
        readonly object lockObject = new object();
        private object dictLock = new object(); // 定义字典的锁对象
        public class Chip
        {
            public double Dimension1 { get; }
            public double Dimension2 { get; }
            public double Dimension3 { get; }
            public double Dimension4 { get; }

            public Chip(double d1, double d2, double d3, double d4)
            {
                Dimension1 = d1;
                Dimension2 = d2;
                Dimension3 = d3;
                Dimension4 = d4;
            }

            public double GetDimensionValue(int dimensionIndex)
            {
                return dimensionIndex switch
                {
                    0 => Dimension1,
                    1 => Dimension2,
                    2 => Dimension3,
                    3 => Dimension4,
                    _ => throw new ArgumentOutOfRangeException(nameof(dimensionIndex), "Invalid dimension index")
                };
            }
        }

        public class Wafer
        {
            public string WaferId { get; set; }
            public List<Chip> Chips { get; set; }

            public Wafer(string waferId)
            {
                WaferId = waferId;
                Chips = new List<Chip>();
            }

            public void AddChip(Chip chip)
            {
                Chips.Add(chip);
            }

            public int GetChipCount()
            {
                return Chips.Count;
            }
        }
        private void LoadTextBoxContent()
        {
            try
            {
                string json = File.ReadAllText("config.json");

                TextBoxContent content = JsonConvert.DeserializeObject<TextBoxContent>(json);

                BinName.Text = content.BinName;
                BinSetting.Text = content.BinSetting;
                dimensionTextBox.Text = content.Dimension;
                filePath.Text = content.FilePath;
                filenameSuffix.Text = content.filenameSuffix;
                para1.Text = content.Para1;
                para1min.Text = content.Para1Min;
                para1rta.Text = content.Para1Rta;
                para1num.Text = content.Para1Num;
                fix1num.Text = content.Fix1num;
                para2.Text = content.Para2;
                para2min.Text = content.Para2Min;
                para2rta.Text = content.Para2Rta;
                para2num.Text = content.Para2Num;
                fix2num.Text = content.Fix2num;
                para3.Text = content.Para3;
                para3min.Text = content.Para3Min;
                para3rta.Text = content.Para3Rta;
                para3num.Text = content.Para3Num;
                fix3num.Text = content.Fix3num;
                para4.Text = content.Para4;
                para4min.Text = content.Para4Min;
                para4rta.Text = content.Para4Rta;
                para4num.Text = content.Para4Num;
                fix4num.Text = content.Fix4num;
            }
            catch (Exception ex)
            {
            }
        }

        private void Page_Unloaded(object sender, RoutedEventArgs e)
        {
            SaveTextBoxContent();
        }

        private void SaveTextBoxContent()
        {
            try
            {
                TextBoxContent content = new TextBoxContent
                {
                    BinName = BinName.Text,
                    BinSetting = BinSetting.Text,
                    Dimension = dimensionTextBox.Text,
                    FilePath = filePath.Text,
                    filenameSuffix = filenameSuffix.Text,

                    Para1 = para1.Text,
                    Para1Min = para1min.Text,
                    Para1Rta = para1rta.Text,
                    Para1Num = para1num.Text,
                    Fix1num = fix1num.Text,

                    Para2 = para2.Text,
                    Para2Min = para2min.Text,
                    Para2Rta = para2rta.Text,
                    Para2Num = para2num.Text,
                    Fix2num = fix2num.Text,

                    Para3 = para3.Text,
                    Para3Min = para3min.Text,
                    Para3Rta = para3rta.Text,
                    Para3Num = para3num.Text,
                    Fix3num = fix3num.Text,

                    Para4 = para4.Text,
                    Para4Min = para4min.Text,
                    Para4Rta = para4rta.Text,
                    Para4Num = para4num.Text,
                    Fix4num = fix4num.Text,
                };

                string json = JsonConvert.SerializeObject(content, Formatting.Indented);
                File.WriteAllText("config.json", json);
            }
            catch (Exception ex)
            {
            }
        }


        public Page1()
        {
            InitializeComponent();
            this.KeepAlive = true;
            DataContext = this;
            LoadTextBoxContent();
            this.Unloaded += Page_Unloaded; // 订阅 Unloaded 事件
            Application.Current.Exit += Application_Exit; // 订阅 Exit 事件
        }

        private void Application_Exit(object sender, ExitEventArgs e)
        {
            SaveTextBoxContent();
        }

        void WriteMatrix(List<double[]>[] pairs, int dim, string output_excel_file)
        {

            // 处理一维数据
            if (dim == 1)
            {
                PrintAndWrite1DMatrix(matrix1D, pairs, output_excel_file);
            }
            else if (dim == 2)
            {
                // 处理二维数据
                PrintAndWrite2DMatrix(matrix2D, pairs, output_excel_file);
            }
            else if (dim == 3)
            {
                // 处理三维数据
                PrintAndWrite3DMatrix(matrix3D, pairs, output_excel_file);
            }
            else if (dim == 4)
            {
                // 处理四维数据
                PrintAndWrite4DMatrix(matrix4D, pairs, output_excel_file);
            }
        }

        bool breakFlag = false;

        Dictionary<string, int> dict2 = new Dictionary<string, int>
        {
            {"TEST", 0}, {"BIN", 1}, {"VF1", 2}, {"VF2", 3}, {"VF3", 4}, {"VF4", 5}, {"VF5", 6}, {"VF6", 7}, {"DVF", 8},
            {"VF", 9}, {"VFD", 10}, {"VZ1", 11}, {"VZ2", 12}, {"IR", 13}, {"LOP1", 14}, {"LOP2", 15}, {"LOP3", 16},
            {"WLP1", 17}, {"WLD1", 18}, {"WLC1", 19}, {"HW1", 20}, {"PURITY1", 21}, {"X1", 22}, {"Y1", 23}, {"Z1", 24},
            {"ST1", 25}, {"INT1", 26}, {"WLP2", 27}, {"WLD2", 28}, {"WLC2", 29}, {"HW2", 30}, {"PURITY2", 31}, {"DVF1", 32},
            {"DVF2", 33}, {"INT2", 34}, {"ST2", 35}, {"VF7", 36}, {"VF8", 37}, {"IR3", 38}, {"IR4", 39}, {"IR5", 40}, {"IR6", 41},
            {"VZ3", 42}, {"VZ4", 43}, {"VZ5", 44}, {"IF", 45}, {"IF1", 46}, {"IF2", 47}, {"ESD1", 48}, {"ESD2", 49}, {"IR1", 50},
            {"IR2", 51}, {"ESD1PASS", 52}, {"ESD2PASS", 53}, {"PosX", 54}, {"PosY", 55}
        };

        private Dictionary<string, Wafer> waferList = new Dictionary<string, Wafer>();

        private async void importWaferFiles(string filename, int dim, int[] col2, string outputCsvFile)
        {
            int flag = 0;
            string outputPath = System.IO.Path.GetDirectoryName(outputCsvFile);
            Wafer waferData = new Wafer(System.IO.Path.GetFileNameWithoutExtension(filename));
            try
            {
                using (StreamReader reader = new StreamReader(filename))
                {

                    while (!reader.EndOfStream)
                    {
                        string[] values = reader.ReadLine().Split(',');
                        if (values.Length >= 3 && values[0] == "TEST" && values[1] == "BIN1" && values[2] == "BIN2")
                        {
                            flag = 1;
                        }

                        string firstValue = values[0];
                        bool isFirstValueAllDigits = Regex.IsMatch(firstValue, @"^\d+$");
                        if (isFirstValueAllDigits && values.Length >= 56)
                        {
                            double Dimension1 = 0;
                            double Dimension2 = 0;
                            double Dimension3 = 0;
                            double Dimension4 = 0;
                            if (dim >= 1)
                            {
                                Dimension1 = !string.IsNullOrEmpty(values[col2[0] + flag]) ? Convert.ToDouble(values[col2[0] + flag]) : -100000;
                            }
                            if (dim >= 2)
                            {
                                Dimension2 = !string.IsNullOrEmpty(values[col2[1] + flag]) ? Convert.ToDouble(values[col2[1] + flag]) : -100000;
                            }
                            if (dim >= 3)
                            {
                                Dimension3 = !string.IsNullOrEmpty(values[col2[2] + flag]) ? Convert.ToDouble(values[col2[2] + flag]) : -100000;
                            }
                            if (dim >= 4)
                            {
                                Dimension4 = !string.IsNullOrEmpty(values[col2[3] + flag]) ? Convert.ToDouble(values[col2[3] + flag]) : -100000;
                            }
                            Chip chipData = new Chip(Dimension1, Dimension2, Dimension3, Dimension4);
                            lock (lockObject)
                            {
                                waferData.Chips.Add(chipData);
                            }
                        }
                    }
                }

                await Task.Run(() =>
                {
                    lock (dictLock)
                    {
                        waferList.Add(waferData.WaferId, waferData);
                    }
                });

                await Dispatcher.InvokeAsync(() =>
                {
                    lock (parameterlockObject)
                    {
                        parameterListBox.Items.Add(System.IO.Path.GetFileName(filename) + " 导入完成!");
                        // 滚动到最新项
                        parameterListBox.ScrollIntoView(parameterListBox.Items[parameterListBox.Items.Count - 1]);
                    }
                });
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"读取文件时出错: {ex.Message}\n{ex.StackTrace}");
                await Task.Run(() =>
                {
                    lock (lockObject)
                    {
                        using (StreamWriter sw = new StreamWriter(System.IO.Path.Combine(outputPath, "NotFindFile.csv"), true, Encoding.UTF8))
                        {
                            sw.WriteLineAsync(filename);
                        }
                    }
                });
                return;
            }
        }

        private async void ProcessFile(Wafer wafer, int dim, List<double[]>[] pairs, double[] fixNums, string[] ops)
        {
            if (waferList.Any())
            {

                if (dim == 1)
                {
                    // 处理一维数据
                    lock (lockObject)
                    {
                        Generate1DMatrix(wafer.Chips, pairs, fixNums, ops);
                    }
                }
                else if (dim == 2)
                {
                    lock (lockObject) 
                    { 
                        // 处理二维数据
                        Generate2DMatrix(wafer.Chips, pairs, fixNums, ops);
                    }
                }
                else if (dim == 3)
                {
                    lock (lockObject)
                    { 
                        // 处理三维数据
                        Generate3DMatrix(wafer.Chips, pairs, fixNums, ops);
                    }
                }
                else if (dim == 4)
                {
                    lock (lockObject)
                    {
                        // 处理四维数据
                        Generate4DMatrix(wafer.Chips, pairs, fixNums, ops);
                    }
                }
            }
        }

        int[] matrix1D;
        int[,] matrix2D;
        int[,,] matrix3D;
        int[,,,] matrix4D;
        public Wafer GetWaferById(string waferId)
        {
            lock (dictLock)
            {
                if (waferList.TryGetValue(waferId, out Wafer wafer))
                {
                    return wafer;
                }
                return null;
            }
        }
        private async void importButton_Click(object sender, RoutedEventArgs e)
        {
            DisableAllButtons();
            waferList.Clear();

            if (!int.TryParse(dimensionTextBox.Text, out int dim))
            {
                System.Windows.MessageBox.Show("Invalid dimension value. Please enter a valid integer.");
                EnableAllButtons();
                return;
            }

            col2 = new int[dim];

            if (dim >= 1)
            {
                if (dict2.ContainsKey(this.para1.Text.ToUpper())) col2[0] = dict2[this.para1.Text.ToUpper()];
            }
            if (dim >= 2)
            {
                if (dict2.ContainsKey(this.para2.Text.ToUpper())) col2[1] = dict2[this.para2.Text.ToUpper()];
            }
            if (dim >= 3)
            {
                if (dict2.ContainsKey(this.para3.Text.ToUpper())) col2[2] = dict2[this.para3.Text.ToUpper()];
            }
            if (dim >= 4)
            {
                MessageBox.Show("4维数据产出分布未开发！");
                EnableAllButtons();
                return;
                if (dict2.ContainsKey(this.para4.Text.ToUpper())) col2[3] = dict2[this.para4.Text.ToUpper()];
            }

            string outputFolder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutputFolder");
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "TXT files (*.txt)|*.txt|All files (*.*)|*.*";
            string output_excel_file = System.IO.Path.Combine(outputFolder, $"{BinName.Text}.xlsx");
            // 检查文件夹是否存在
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            string output_csv_file = System.IO.Path.Combine(outputFolder, "每片颗粒数.csv");
            string not_find_csv_file = System.IO.Path.Combine(outputFolder, "NotFindFile.csv");
            try
            {
                if (File.Exists(output_csv_file))
                {
                    File.Delete(output_csv_file);
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show($"Error deleting file {output_csv_file}: {ex.Message}");
                EnableAllButtons();
                return;
            }

            try
            {
                if (File.Exists(not_find_csv_file))
                {
                    File.Delete(not_find_csv_file);
                }
            }
            catch (IOException ex)
            {
                MessageBox.Show($"Error deleting file {not_find_csv_file}: {ex.Message}");
                EnableAllButtons();
                return;
            }

            string filePathText = filePath.Text;

            string[] lines = multiLineTextBox.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            int totalLines = lines.Count();
            DateTime startTime = DateTime.Now; // 记录开始时间
            List<Task> tasks = new List<Task>(); // 声明 tasks 列表
            Progress = 0;
            progressBar.Value = 0;
            int processedLines = 0;
            string fileSuffix = this.filenameSuffix.Text;
            // 逐行读取并处理
            foreach (string line in lines)
            {

                string filePathTemp = System.IO.Path.Combine(filePathText, line + fileSuffix + ".csv");
                tasks.Add(Task.Run(() =>
                {
                    importWaferFiles(filePathTemp, dim, col2, output_csv_file);
                    processedLines++;
                    Progress = (int)Math.Ceiling(processedLines * 100.0 / totalLines);
                    Dispatcher.Invoke(() =>
                    {
                        progressBar.Value = Progress;
                        progressText.Text = $"{Progress}%";
                    });
                }));

                await Task.WhenAll(tasks); // 等待所有任务完成

                foreach (var temp in waferList)
                {
                    tasks.Add(Task.Run(() =>
                    {
                        var wafer = GetWaferById(temp.Key);

                        if (wafer != null)
                        {
                            lock (fileWriteLock) // 使用锁保护写入操作
                            {
                                using (var sw = new StreamWriter(output_csv_file, true, Encoding.UTF8))
                                {
                                    sw.WriteLine(temp.Key + "," + wafer.Chips.Count);
                                }
                            }
                        }
                    }));
                }
                await Task.WhenAll(tasks); // 等待所有任务完成
            }
            MessageBox.Show("导入文件成功！");

            EnableAllButtons();
        }

        int[] col2;

        private double CalculateNewValue(double currentValue, string operation, double value)
        {
            switch (operation)
            {
                case "加":
                    return currentValue + value;
                case "减":
                    return currentValue - value;
                case "乘":
                    return currentValue * value;
                case "除":
                    if (value == 0)
                    {
                        MessageBox.Show("Cannot divide by zero.");
                        return currentValue;
                    }
                    return currentValue / value;
                default:
                    return currentValue;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
        private void DisableAllButtons()
        {
            foreach (var control in MainGrid.Children)
            {
                if (control is Button button)
                {
                    button.IsEnabled = false;
                }
            }
        }

        private void EnableAllButtons()
        {
            foreach (var control in MainGrid.Children)
            {
                if (control is Button button)
                {
                    button.IsEnabled = true;
                }
            }
        }

        public void ExportToCustomJson(Dictionary<string, List<double[]>> pairsDict, string filePath)
        {
            var sb = new StringBuilder();
            sb.AppendLine("{");

            foreach (var pair in pairsDict)
            {
                sb.AppendLine($"  \"{pair.Key}\": [");

                for (int i = 0; i < pair.Value.Count; i++)
                {
                    var pairArray = pair.Value[i];
                    sb.Append("    [");
                    sb.Append(string.Join(", ", pairArray));
                    sb.Append("]");

                    if (i < pair.Value.Count - 1)
                    {
                        sb.AppendLine(",");
                    }
                    else
                    {
                        sb.AppendLine();
                    }
                }

                sb.AppendLine("  ],");
            }

            // Remove the trailing comma and newline
            if (sb.Length > 2)
            {
                sb.Length -= 3;
            }

            sb.AppendLine("}");

            File.WriteAllText(filePath, sb.ToString());
        }

        private async void runButton_Click(object sender, RoutedEventArgs e)
        {
            DisableAllButtons();
            Progress = 0;

            string filePath = BinName.Text;

            if (!int.TryParse(dimensionTextBox.Text, out int dim))
            {
                System.Windows.MessageBox.Show("Invalid dimension value. Please enter a valid integer.");
                EnableAllButtons();
                return;
            }

            // 定义字典
            Dictionary<string, int> dict = new Dictionary<string, int>
                {
                    {"VF1", 4}, {"VF2", 6}, {"VF3", 8}, {"VF4", 10}, {"VZ1", 12},
                    {"IR", 14}, {"HW1", 16}, {"LOP1", 18}, {"WLP1", 20}, {"WLD1", 22},
                    {"IR1", 24}, {"VFD", 26}, {"DVF", 28}, {"IR2", 30}, {"WLC1", 32},
                    {"VF5", 34}, {"VF6", 36}, {"VF7", 38}, {"VF8", 40}, {"DVF1", 42},
                    {"DVF2", 44}, {"VZ2", 46}, {"VZ3", 48}, {"VZ4", 50}, {"VZ5", 52},
                    {"IR3", 54}, {"IR4", 56}, {"IR5", 58}, {"IR6", 60}, {"IF", 62},
                    {"IF1", 64}, {"IF2", 66}, {"LOP2", 68}, {"WLP2", 70}, {"WLD2", 72},
                    {"HW2", 74}, {"WLC2", 76}
                };


            int dimension = dim;
            if (dimension < 1 || dimension > 4)
            {
                MessageBox.Show("Dimension must be between 1 and 4.");
                EnableAllButtons();
                return;
            }

            double[] minValues = new double[dimension];
            double[] stepSizes = new double[dimension];
            double[] fixNums = new double[dimension];
            int[] counts = new int[dimension];
            int[] col = new int[dimension];
            string[] paraStrings = new string[dimension];
            string[] ops = new string[dimension];
            string output_excel_file = "";
            // 处理并输出结果
            string fileExtension = ".xlsx";
            string outputFolder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutputFolder");

            if (dimension >= 1)
            {
                if (dict.ContainsKey(this.para1.Text.ToUpper())) col[0] = dict[this.para1.Text.ToUpper()];
                paraStrings[0] = this.para1.Text.ToUpper();
                minValues[0] = double.Parse(this.para1min.Text);
                stepSizes[0] = double.Parse(this.para1rta.Text);
                counts[0] = int.Parse(this.para1num.Text);
                fixNums[0] = Math.Round(double.Parse(this.fix1num.Text),6);
                ops[0] = oprBox1.Text;
                output_excel_file = System.IO.Path.Combine(outputFolder, $"{filePath}_{para1.Text}");
            }

            if (dimension >= 2)
            {
                if (dict.ContainsKey(this.para2.Text.ToUpper())) col[1] = dict[this.para2.Text.ToUpper()];
                paraStrings[1] = this.para2.Text.ToUpper();

                minValues[1] = double.Parse(this.para2min.Text);
                stepSizes[1] = double.Parse(this.para2rta.Text);
                counts[1] = int.Parse(this.para2num.Text);
                fixNums[1] = Math.Round(double.Parse(this.fix2num.Text), 6);
                ops[1] = oprBox2.Text;
                output_excel_file = System.IO.Path.Combine(outputFolder, $"{filePath}_{para1.Text}_{para2.Text}");

            }

            if (dimension >= 3)
            {
                if (dict.ContainsKey(this.para3.Text.ToUpper())) col[2] = dict[this.para3.Text.ToUpper()];
                paraStrings[2] = this.para3.Text.ToUpper();
                minValues[2] = double.Parse(this.para3min.Text);
                stepSizes[2] = double.Parse(this.para3rta.Text);
                counts[2] = int.Parse(this.para3num.Text);
                fixNums[2] = Math.Round(double.Parse(this.fix3num.Text), 6);
                ops[2] = oprBox3.Text;
                output_excel_file = System.IO.Path.Combine(outputFolder, $"{filePath}_{para1.Text}_{para2.Text}_{para3.Text}");

            }

            if (dimension >= 4)
            {
                MessageBox.Show("4维数据产出分布未开发！");
                EnableAllButtons();
                return;
                if (dict.ContainsKey(this.para4.Text.ToUpper())) col[3] = dict[this.para4.Text.ToUpper()];
                paraStrings[3] = this.para4.Text.ToUpper();
                minValues[3] = double.Parse(this.para4min.Text);
                stepSizes[3] = double.Parse(this.para4rta.Text);
                counts[3] = int.Parse(this.para4num.Text);
                fixNums[3] = Math.Round(double.Parse(this.fix4num.Text), 6);
                ops[3] = oprBox4.Text;
                output_excel_file = System.IO.Path.Combine(outputFolder, $"{filePath}_{para1.Text}_{para2.Text}_{para3.Text}_{para4.Text}");
            }

            // 存储生成的pair数组
            List<double[]>[] pairs = new List<double[]>[dimension];

            // 生成pair数组
            for (int i = 0; i < dimension; i++)
            {
                pairs[i] = new List<double[]>();
                for (int j = 0; j < counts[i]; j++)
                {
                    double first = Math.Round(minValues[i] + j * stepSizes[i],3);
                    double second = Math.Round(first + stepSizes[i], 3);
                    pairs[i].Add(new double[] { first, second });
                }
            }

            if (dimension == 1)
            {
                int length = pairs[0].Count;
                matrix1D = new int[length + 2];
            }

            else if (dimension == 2)
            {
                int rows = pairs[0].Count;
                int cols = pairs[1].Count;
                matrix2D = new int[rows + 2, cols + 2];
            }
            else if (dimension == 3)
            {
                int rows = pairs[0].Count;
                int cols = pairs[1].Count;
                int depth = pairs[2].Count;
                matrix3D = new int[rows + 2, cols + 2, depth + 2];
            }
            else if (dimension == 4)
            {
                int rows = pairs[0].Count;
                int cols = pairs[1].Count;
                int depth = pairs[2].Count;
                int time = pairs[3].Count;
                matrix4D = new int[rows + 2, cols + 2, depth + 2, time + 2];
            }

            
            // 检查文件是否存在并生成唯一文件名
            int counter = 1;
            while (File.Exists(output_excel_file + fileExtension))
            {
                string newFileName = $"{output_excel_file}_{counter}";
                output_excel_file = System.IO.Path.Combine(outputFolder, newFileName);
                counter++;
            }

            // 检查文件夹是否存在
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            DateTime startTime = DateTime.Now; // 记录开始时间

            Progress = 0;

            progressBar.Value = 0;

            int processedLines = 0;

            int totalWafers = waferList.Count;

            // 创建文件夹
            List<Task> tasks = new List<Task>(); // 声明 tasks 列表
            output_excel_file = output_excel_file + fileExtension;
            foreach (var wafer in waferList)
            {
                tasks.Add(Task.Run(() =>
                {

                    ProcessFile(wafer.Value, dimension, pairs, fixNums, ops);

                    Dispatcher.Invoke(() =>
                    {
                        processedLines++;
                        Progress = processedLines * 100 / totalWafers;
                        progressBar.Value = Progress;
                        progressText.Text = $"{Progress}%";
                    });

                }));
            }
            // 等待所有任务完成
            await Task.WhenAll(tasks);

            // 写入处理后的数据到 Excel 文件
            WriteMatrix(pairs, dim, output_excel_file);

            if (!breakFlag)
            {
                DateTime endTime = DateTime.Now; // 记录结束时间
                TimeSpan totalTime = endTime - startTime; // 计算运行时间
                                                            // 弹出消息框询问是否打开文件
                MessageBoxResult result = MessageBox.Show("Excel 文件已导出到 " + output_excel_file + $", 总共耗时：{totalTime.TotalSeconds} 秒 , \n是否打开该文件？", "导出成功", MessageBoxButton.YesNo, MessageBoxImage.Question);

                // 根据用户的选择执行相应的操作
                if (result == MessageBoxResult.Yes)
                {
                    // 打开文件
                    try
                    {
                        Process.Start(new ProcessStartInfo { FileName = output_excel_file, UseShellExecute = true });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("打开文件失败：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    // 处理文件名为空的情况
                    MessageBox.Show("Excel 文件: " + output_excel_file + "导出出错！");
                }
            }
            EnableAllButtons();
        }

        private async void runButton_Click_New(object sender, RoutedEventArgs e)
        {
            DisableAllButtons();
            Progress = 0;

            string filePath = BinName.Text;

            if (!int.TryParse(dimensionTextBox.Text, out int dim))
            {
                System.Windows.MessageBox.Show("Invalid dimension value. Please enter a valid integer.");
                EnableAllButtons();
                return;
            }

            string outputFolder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutputFolder");
            string JsonfilePath = System.IO.Path.Combine(outputFolder, "pairsConfig.json"); ;

            // 定义字典
            Dictionary<string, int> dict = new Dictionary<string, int>
            {
                {"VF1", 4}, {"VF2", 6}, {"VF3", 8}, {"VF4", 10}, {"VZ1", 12},
                {"IR", 14}, {"HW1", 16}, {"LOP1", 18}, {"WLP1", 20}, {"WLD1", 22},
                {"IR1", 24}, {"VFD", 26}, {"DVF", 28}, {"IR2", 30}, {"WLC1", 32},
                {"VF5", 34}, {"VF6", 36}, {"VF7", 38}, {"VF8", 40}, {"DVF1", 42},
                {"DVF2", 44}, {"VZ2", 46}, {"VZ3", 48}, {"VZ4", 50}, {"VZ5", 52},
                {"IR3", 54}, {"IR4", 56}, {"IR5", 58}, {"IR6", 60}, {"IF", 62},
                {"IF1", 64}, {"IF2", 66}, {"LOP2", 68}, {"WLP2", 70}, {"WLD2", 72},
                {"HW2", 74}, {"WLC2", 76}
            };




            int dimension = dim;

            if (dimension < 1 || dimension > 4)
            {
                MessageBox.Show("Dimension must be between 1 and 4.");
                EnableAllButtons();
                return;
            }
            // 读取 JSON 文件内容
            string jsonContent = File.ReadAllText(JsonfilePath);

            // 解析 JSON 为字典
            var pairsDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, List<double[]>>>(jsonContent);

            if (pairsDict.Count != dimension)
            {
                MessageBox.Show("Json Dimension is not right!");
                EnableAllButtons();
                return;
            }


            double[] fixNums = new double[dimension];
            int[] col = new int[dimension];
            string[] paraStrings = new string[dimension];
            string[] ops = new string[dimension];
            string output_excel_file = "";
            // 处理并输出结果
            string fileExtension = ".xlsx";


            if (dimension >= 1)
            {
                if (dict.ContainsKey(this.para1.Text.ToUpper())) col[0] = dict[this.para1.Text.ToUpper()];
                paraStrings[0] = this.para1.Text.ToUpper();
                fixNums[0] = Math.Round(double.Parse(this.fix1num.Text), 6);
                ops[0] = oprBox1.Text;
                output_excel_file = System.IO.Path.Combine(outputFolder, $"{filePath}_{para1.Text}");
            }

            if (dimension >= 2)
            {
                if (dict.ContainsKey(this.para2.Text.ToUpper())) col[1] = dict[this.para2.Text.ToUpper()];
                paraStrings[1] = this.para2.Text.ToUpper();
                fixNums[1] = Math.Round(double.Parse(this.fix2num.Text), 6);
                ops[1] = oprBox2.Text;
                output_excel_file = System.IO.Path.Combine(outputFolder, $"{filePath}_{para1.Text}_{para2.Text}");

            }

            if (dimension >= 3)
            {
                if (dict.ContainsKey(this.para3.Text.ToUpper())) col[2] = dict[this.para3.Text.ToUpper()];
                paraStrings[2] = this.para3.Text.ToUpper();
                fixNums[2] = Math.Round(double.Parse(this.fix3num.Text), 6);
                ops[2] = oprBox3.Text;
                output_excel_file = System.IO.Path.Combine(outputFolder, $"{filePath}_{para1.Text}_{para2.Text}_{para3.Text}");

            }

            if (dimension >= 4)
            {
                MessageBox.Show("4维数据产出分布未开发！");
                EnableAllButtons();
                return;
                if (dict.ContainsKey(this.para4.Text.ToUpper())) col[3] = dict[this.para4.Text.ToUpper()];
                paraStrings[3] = this.para4.Text.ToUpper();
                fixNums[3] = Math.Round(double.Parse(this.fix4num.Text), 6);
                ops[3] = oprBox4.Text;
                output_excel_file = System.IO.Path.Combine(outputFolder, $"{filePath}_{para1.Text}_{para2.Text}_{para3.Text}_{para4.Text}");
            }

            // 存储生成的pair数组
            List<double[]>[] pairs = new List<double[]>[dimension];
            int i = 0;
            foreach (var pair in pairsDict)
            {
                pairs[i] = new List<double[]>();
                if(pair.Key.Trim() != paraStrings[i].Trim())
                {
                    MessageBox.Show("Json Para is not correct!");
                    EnableAllButtons();
                    return;
                }

                foreach (var subArray in pair.Value)
                {
                    pairs[i].Add(subArray);
                }
                i++;
            }

            if (dimension == 1)
            {
                int length = pairs[0].Count;
                matrix1D = new int[length + 2];
            }

            else if (dimension == 2)
            {
                int rows = pairs[0].Count;
                int cols = pairs[1].Count;
                matrix2D = new int[rows + 2, cols + 2];
            }
            else if (dimension == 3)
            {
                int rows = pairs[0].Count;
                int cols = pairs[1].Count;
                int depth = pairs[2].Count;
                matrix3D = new int[rows + 2, cols + 2, depth + 2];
            }
            else if (dimension == 4)
            {
                int rows = pairs[0].Count;
                int cols = pairs[1].Count;
                int depth = pairs[2].Count;
                int time = pairs[3].Count;
                matrix4D = new int[rows + 2, cols + 2, depth + 2, time + 2];
            }


            // 检查文件是否存在并生成唯一文件名
            int counter = 1;
            while (File.Exists(output_excel_file + fileExtension))
            {
                string newFileName = $"{output_excel_file}_{counter}";
                output_excel_file = System.IO.Path.Combine(outputFolder, newFileName);
                counter++;
            }

            // 检查文件夹是否存在
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            DateTime startTime = DateTime.Now; // 记录开始时间

            Progress = 0;

            progressBar.Value = 0;

            int processedLines = 0;

            int totalWafers = waferList.Count;

            // 创建文件夹
            List<Task> tasks = new List<Task>(); // 声明 tasks 列表
            output_excel_file = output_excel_file + fileExtension;
            foreach (var wafer in waferList)
            {
                tasks.Add(Task.Run(() =>
                {

                    ProcessFile(wafer.Value, dimension, pairs, fixNums, ops);

                    Dispatcher.Invoke(() =>
                    {
                        processedLines++;
                        Progress = processedLines * 100 / totalWafers;
                        progressBar.Value = Progress;
                        progressText.Text = $"{Progress}%";
                    });

                }));
            }
            // 等待所有任务完成
            await Task.WhenAll(tasks);

            // 写入处理后的数据到 Excel 文件
            WriteMatrix(pairs, dim, output_excel_file);

            if (!breakFlag)
            {
                DateTime endTime = DateTime.Now; // 记录结束时间
                TimeSpan totalTime = endTime - startTime; // 计算运行时间
                                                          // 弹出消息框询问是否打开文件
                MessageBoxResult result = MessageBox.Show("Excel 文件已导出到 " + output_excel_file + $", 总共耗时：{totalTime.TotalSeconds} 秒 , \n是否打开该文件？", "导出成功", MessageBoxButton.YesNo, MessageBoxImage.Question);

                // 根据用户的选择执行相应的操作
                if (result == MessageBoxResult.Yes)
                {
                    // 打开文件
                    try
                    {
                        Process.Start(new ProcessStartInfo { FileName = output_excel_file, UseShellExecute = true });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("打开文件失败：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    // 处理文件名为空的情况
                    MessageBox.Show("Excel 文件: " + output_excel_file + "导出出错！");
                }
            }
            EnableAllButtons();
        }
        void Generate1DMatrix(List<Chip> chips, List<double[]>[] pairs, double[]fixNums,string[] ops)
        {
            // 执行计算或处理任务
            foreach (var chip in chips)
            {
                int index = GetRangeIndex(CalculateNewValue(chip.GetDimensionValue(0), ops[0], fixNums[0]), pairs[0]);

                if (index >= 0)
                {
                    matrix1D[index]++;
                }
            }
        }

        void Generate2DMatrix(List<Chip> chips, List<double[]>[] pairs, double[] fixNums, string[] ops)
        {
            foreach(var chip in chips)
            {
                int rowIndex = GetRangeIndex(CalculateNewValue(chip.GetDimensionValue(0), ops[0], fixNums[0]), pairs[0]);
                int colIndex = GetRangeIndex(CalculateNewValue(chip.GetDimensionValue(1), ops[1], fixNums[1]), pairs[1]);

                if (rowIndex >= 0 && colIndex >= 0)
                {
                    matrix2D[rowIndex, colIndex]++;
                }
            }

        }

        void Generate3DMatrix(List<Chip> chips, List<double[]>[] pairs, double[] fixNums, string[] ops)
        {
            // 执行计算或处理任务
            foreach (var chip in chips)
            {
                int rowIndex = GetRangeIndex(CalculateNewValue(chip.GetDimensionValue(0), ops[0], fixNums[0]), pairs[0]);
                int colIndex = GetRangeIndex(CalculateNewValue(chip.GetDimensionValue(1), ops[1], fixNums[1]), pairs[1]);
                int depthIndex = GetRangeIndex(CalculateNewValue(chip.GetDimensionValue(2), ops[2], fixNums[2]), pairs[2]);

                if (rowIndex >= 0 && colIndex >= 0 && depthIndex >= 0)
                {
                    matrix3D[rowIndex, colIndex, depthIndex]++;
                }
            }
        }

        void Generate4DMatrix(List<Chip> chips, List<double[]>[] pairs, double[] fixNums, string[] ops)
        {            // 执行计算或处理任务
            foreach (var chip in chips)
            {
                int rowIndex = GetRangeIndex(CalculateNewValue(chip.GetDimensionValue(0), ops[0], fixNums[0]), pairs[0]);
                int colIndex = GetRangeIndex(CalculateNewValue(chip.GetDimensionValue(1), ops[1], fixNums[1]), pairs[1]);
                int depthIndex = GetRangeIndex(CalculateNewValue(chip.GetDimensionValue(2), ops[2], fixNums[2]), pairs[2]);
                int timeIndex = GetRangeIndex(CalculateNewValue(chip.GetDimensionValue(3), ops[3], fixNums[3]), pairs[3]);

                if (rowIndex >= 0 && colIndex >= 0 && depthIndex >= 0 && timeIndex >= 0)
                {
                    matrix4D[rowIndex, colIndex, depthIndex, timeIndex]++;
                }
            }
        }
        int GetRangeIndex(double value, List<double[]> pairs)
        {
            for (int i = 0; i < pairs.Count; i++)
            {
                if (value >= pairs[i][0] && value < pairs[i][1])
                {
                    return i + 1;
                }else if(value < pairs[i][0])
                {
                    return 0;
                }else if(value >= pairs[pairs.Count - 1][1])
                {
                    return pairs.Count+1;
                }
            }
            return -1;  // 不在任何范围内
        }
        void PrintAndWrite1DMatrix(int[] matrix, List<double[]>[] pairs, string output_excel_file)
        {
            string p1 = this.para1.Text.ToUpper();

            int rows = matrix.GetLength(0);

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage();
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("matrix1D");

            worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            int row1 = 1;
            worksheet.Cells[row1, 1].Value = BinName.Text;
            worksheet.Cells[row1, 2].Value = "片量";
            worksheet.Cells[row1, 3].Value = waferList.Count;
            worksheet.Cells[row1 + 1, 2].Value = "总计百分比";
            
            double rowstotal = 0;
            double total = 0;
            worksheet.Cells[1 + row1, 1].Value = ($"{p1}");

            for (int i = 0; i < rows; i++)
            {
                rowstotal += matrix[i];
            }
            total = rowstotal;

            for (int i = 0; i < rows; i++)
            {

                if (i == 0)
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($"<{pairs[0][0][0]}");
                }
                else if (i == rows - 1)
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($">{pairs[0][i - 2][1]}");
                }
                else
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($"{pairs[0][i - 1][0]}-{pairs[0][i - 1][1]}");
                }

                worksheet.Cells[row1 + 2 + i, 2].Value = (matrix[i]) / total;
                worksheet.Cells[row1 + 2 + i, 2].Style.Numberformat.Format = "0.00%";
            }
            worksheet.Cells[row1 + 2 + rows, 1].Value = "Sum";
            worksheet.Cells[row1 + 2 + rows, 2].Value = rowstotal / total;
            worksheet.Cells[row1 + 2 + rows, 2].Style.Numberformat.Format = "0.00%";
            row1 = row1 + 2 + rows + 2;
            worksheet.Cells[row1, 1].Value = BinName.Text;
            worksheet.Cells[row1, 2].Value = "片量";
            worksheet.Cells[row1, 3].Value = waferList.Count;
            worksheet.Cells[row1 +1 , 2].Value = "颗粒数";

            worksheet.Cells[1 + row1, 1].Value = ($"{p1}");

            for (int i = 0; i < rows; i++)
            {

                if (i == 0)
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($"<{pairs[0][0][0]}");
                }
                else if (i == rows - 1)
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($">{pairs[0][i - 2][1]}");
                }
                else
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($"{pairs[0][i - 1][0]}-{pairs[0][i - 1][1]}");
                }

                worksheet.Cells[row1 + 2 + i, 2].Value = (matrix[i]) ;
            }
            worksheet.Cells[row1 + 2 + rows, 1].Value = "Sum";
            worksheet.Cells[row1 + 2 + rows, 2].Value = rowstotal ;

            int row2 = row1 + 2 + rows;
            int col2 = 2;
            // 设置第一行的字体为微软雅黑、大小为14号
            // 获取单元格范围对象
            var address = new ExcelAddress(2, 1, row2, col2);
            
            using (ExcelRange range = worksheet.Cells[address.Address])
            {
                range.Style.Font.Name = "微软雅黑";
                range.Style.Font.Size = 11;
            }

            worksheet.Cells.AutoFitColumns();

            // 确保文件名不为空
            if (!string.IsNullOrEmpty(output_excel_file))
            {
                FileInfo excelFile = new FileInfo(output_excel_file);

                // 检查文件是否已存在
                if (excelFile.Exists)
                {
                    try
                    {
                        // 文件未被打开，继续保存 Excel 文件
                        excelPackage.SaveAs(excelFile);
                    }
                    catch (IOException)
                    {
                        // Excel 文件已存在并且被打开，弹出提示框显示
                        MessageBox.Show($"文件 {output_excel_file} 已存在并且被打开，请关闭后重新保存。", "文件已打开", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                        EnableAllButtons();
                        return;
                    }
                }
                else
                {
                    // 文件不存在，直接保存 Excel 文件
                    excelPackage.SaveAs(excelFile);
                }
            }
        }
        
        public void CreateExcelWithColorScale(ExcelWorksheet worksheet, int row1, int col1,int row2,int col2)
        {

            // Define the address of the cells to which we will apply the conditional formatting
            var address = new ExcelAddress(row1, col1, row2, col2);

            // Add a new conditional formatting rule for color scale
            var conditionalFormatting = worksheet.ConditionalFormatting.AddThreeColorScale(address);

            // Define minimum, midpoint, and maximum values with corresponding colors
            conditionalFormatting.LowValue.Type = eExcelConditionalFormattingValueObjectType.Min;
            conditionalFormatting.LowValue.Color = ColorTranslator.FromHtml("#63BE7B"); // Green

            conditionalFormatting.MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Percentile;
            conditionalFormatting.MiddleValue.Value = 50;
            conditionalFormatting.MiddleValue.Color = ColorTranslator.FromHtml("#FFEB84"); // Yellow

            conditionalFormatting.HighValue.Type = eExcelConditionalFormattingValueObjectType.Max;
            conditionalFormatting.HighValue.Color = ColorTranslator.FromHtml("#F8696B"); // Red

            // 获取单元格范围对象
            address = new ExcelAddress(row1, col1, row2+1, col2);
            ExcelRange range = worksheet.Cells[address.Address];

            // 设置边框样式
            range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        }

        void PrintAndWrite2DMatrix(int[,] matrix, List<double[]>[] pairs, string output_excel_file)
        {
            string p1 = this.para1.Text.ToUpper();
            string p2 = this.para2.Text.ToUpper();

            int rows = matrix.GetLength(0);
            int cols = matrix.GetLength(1);

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage();
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("matrix2D");

            worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            double matrixTotal = 0;
            double total = 0;
            double[] rowTotals = new double[rows];

            for (int i = 0; i < rows; i++)
            {
                double rowtotal = 0;
                for (int j = 0; j < cols; j++)
                {
                    rowtotal += matrix[i, j];
                    matrixTotal += matrix[i, j];
                }
                rowTotals[i] = rowtotal;
            }
            total += matrixTotal;


            int row1 = 1;

            worksheet.Cells[row1, 1].Value = BinName.Text;
            worksheet.Cells[row1, 2].Value = "片量";
            worksheet.Cells[row1, 3].Value = waferList.Count;
            worksheet.Cells[row1, 8].Value = "总计百分比";
            // 获取单元格范围对象
            var address = new ExcelAddress(row1, 1, row1, 8);
            ExcelRange range = worksheet.Cells[address.Address];
            range.Style.Font.Bold = true;
            range.Style.Font.Color.SetColor(System.Drawing.Color.Red);

            worksheet.Cells[row1 + 1, 1].Value = ($"{p1}/{p2}");

            for (int j = 0; j < cols; j++)
            {
                if (j == 0)
                {
                    worksheet.Cells[row1 + 1, j + 2].Value = ($"<{pairs[1][0][0]}");
                }
                else if (j == cols - 1)
                {
                    worksheet.Cells[row1 + 1, j + 2].Value = ($">{pairs[1][j -2][1]}");
                }
                else
                {
                    worksheet.Cells[row1 + 1, j + 2].Value = ($"{pairs[1][j - 1][0]}-{pairs[1][j - 1][1]}");
                }
            }
            worksheet.Cells[row1 + 1, cols + 2].Value = "Sum";
            for (int i = 0; i < rows; i++)
            {
                double rowstotal = 0;

                if (i == 0)
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($"<{pairs[0][0][0]}");
                }
                else if (i == rows - 1)
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($">{pairs[0][i - 2][1]}");
                }
                else
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($"{pairs[0][i - 1][0]}-{pairs[0][i - 1][1]}");
                }
                for (int j = 0; j < cols; j++)
                {
                    rowstotal += matrix[i, j];
                    worksheet.Cells[row1 + 1 + 1 + i, j + 2].Value = (matrix[i, j]) / total;
                    worksheet.Cells[row1 + 1 + 1 + i, j + 2].Style.Numberformat.Format = "0.00%";
                }
                worksheet.Cells[row1 + 1 + 1 + i, cols + 2].Value = rowstotal / total;
                worksheet.Cells[row1 + 1 + 1 + i, cols + 2].Style.Numberformat.Format = "0.00%";
            }
            worksheet.Cells[row1 + 1 + 1 + rows, 1].Value = "Sum";

            for (int j = 0; j < cols; j++)
            {
                double colstotal = 0;

                for (int i = 0; i < rows; i++)
                {
                    colstotal += matrix[i, j];
                }
                worksheet.Cells[row1 + 1 + 1 + rows, j + 2].Value = colstotal / total;
                worksheet.Cells[row1 + 1 + 1 + rows, j + 2].Style.Numberformat.Format = "0.00%";
            }
            worksheet.Cells[row1 + 1 + 1 + rows, cols + 2].Value = matrixTotal / total;
            worksheet.Cells[row1 + 1 + 1 + rows, cols + 2].Style.Numberformat.Format = "0.00%";


            CreateExcelWithColorScale(worksheet, row1 + 1 , 1, row1 + 1 + rows, cols + 2);

            // next rows total percent matrix 

            // next rows total percent matrix 

            // next rows total percent matrix 

            row1 = rows + 2 + row1;
            worksheet.Cells[row1 + 2, 1].Value = BinName.Text;
            worksheet.Cells[row1 + 2, 2].Value = "片量";
            worksheet.Cells[row1 + 2, 3].Value = waferList.Count;
            worksheet.Cells[row1 + 2, 8].Value = "行百分比";

            address = new ExcelAddress(row1 + 2, 1, row1 + 2, 8);
            range = worksheet.Cells[address.Address];
            range.Style.Font.Bold = true;
            range.Style.Font.Color.SetColor(System.Drawing.Color.Red);

            row1 += 1;
            worksheet.Cells[row1 + 2, 1].Value = ($"{p1}/{p2}");
            for (int j = 0; j < cols; j++)
            {
                if (j == 0)
                {
                    worksheet.Cells[row1 + 1 + 1, j + 2].Value = ($"<{pairs[1][0][0]}");
                }
                else if (j == cols - 1)
                {
                    worksheet.Cells[row1 + 1 + 1, j + 2].Value = ($">{pairs[1][j - 2][1]}");
                }
                else
                {
                    worksheet.Cells[row1 + 1 + 1, j + 2].Value = ($"{pairs[1][j - 1][0]}-{pairs[1][j - 1][1]}");
                }
            }
            worksheet.Cells[row1 + 1 + 1, cols + 2].Value = "Sum";
            for (int i = 0; i < rows; i++)
            {
                double rowstotal = 0;

                if (i == 0)
                {
                    worksheet.Cells[row1 + 1 + 2 + i, 1].Value = ($"<{pairs[0][0][0]}");
                }
                else if (i == rows - 1)
                {
                    worksheet.Cells[row1 + 1 + 2 + i, 1].Value = ($">{pairs[0][i - 2][1]}");
                }
                else
                {
                    worksheet.Cells[row1 + 1 + 2 + i, 1].Value = ($"{pairs[0][i - 1][0]}-{pairs[0][i - 1][1]}");
                }
                for (int j = 0; j < cols; j++)
                {
                    rowstotal += matrix[i, j];
                    worksheet.Cells[row1 + 1 + 2 + i, j + 2].Value = (matrix[i, j]) / rowTotals[i];
                    worksheet.Cells[row1 + 1 + 2 + i, j + 2].Style.Numberformat.Format = "0.00%";
                }
                worksheet.Cells[row1 + 1 + 2 + i, cols + 2].Value = rowstotal / rowTotals[i];
                worksheet.Cells[row1 + 1 + 2 + i, cols + 2].Style.Numberformat.Format = "0.00%";
            }
            worksheet.Cells[row1 + 1 + 2 + rows, 1].Value = "Sum";

            

            for (int j = 0; j < cols; j++)
            {
                double colstotal = 0;

                for (int i = 0; i < rows; i++)
                {
                    colstotal += matrix[i, j];
                }
                worksheet.Cells[row1 + 1 + 2 + rows, j + 2].Value = colstotal / total;
                worksheet.Cells[row1 + 1 + 2 + rows, j + 2].Style.Numberformat.Format = "0.00%";
            }
            worksheet.Cells[row1 + 1 + 2 + rows, cols + 2].Value = matrixTotal / total;
            worksheet.Cells[row1 + 1 + 2 + rows, cols + 2].Style.Numberformat.Format = "0.00%";

            CreateExcelWithColorScale(worksheet, row1 + 1 + 1, 1, row1 + 1 + 1 + rows, cols + 2);

            // next rows count matrix 
            // next rows count matrix 
            // next rows count matrix 
            // next rows count matrix 
            // next rows count matrix 
            // next rows count matrix 
            row1 = rows + 2 + row1 + 1;
            worksheet.Cells[row1 + 2, 1].Value = BinName.Text;
            worksheet.Cells[row1 + 2, 2].Value = "片量";
            worksheet.Cells[row1 + 2, 3].Value = waferList.Count;
            worksheet.Cells[row1 + 2, 8].Value = "颗粒数分布";

            address = new ExcelAddress(row1 + 2, 1, row1 + 2, 8);
            range = worksheet.Cells[address.Address];
            range.Style.Font.Bold = true;
            range.Style.Font.Color.SetColor(System.Drawing.Color.Red);

            row1 += 1;
            worksheet.Cells[row1 + 2, 1].Value = ($"{p1}/{p2}");
            for (int j = 0; j < cols; j++)
            {
                if (j == 0)
                {
                    worksheet.Cells[row1 + 1 + 1, j + 2].Value = ($"<{pairs[1][0][0]}");
                }
                else if (j == cols - 1)
                {
                    worksheet.Cells[row1 + 1 + 1, j + 2].Value = ($">{pairs[1][j - 2][1]}");
                }
                else
                {
                    worksheet.Cells[row1 + 1 + 1, j + 2].Value = ($"{pairs[1][j - 1][0]}-{pairs[1][j - 1][1]}");
                }
            }
            worksheet.Cells[row1 + 1 + 1, cols + 2].Value = "Sum";
            for (int i = 0; i < rows; i++)
            {
                double rowstotal = 0;

                if (i == 0)
                {
                    worksheet.Cells[row1 + 1 + 2 + i, 1].Value = ($"<{pairs[0][0][0]}");
                }
                else if (i == rows - 1)
                {
                    worksheet.Cells[row1 + 1 + 2 + i, 1].Value = ($">{pairs[0][i - 2][1]}");
                }
                else
                {
                    worksheet.Cells[row1 + 1 + 2 + i, 1].Value = ($"{pairs[0][i - 1][0]}-{pairs[0][i - 1][1]}");
                }
                for (int j = 0; j < cols; j++)
                {
                    rowstotal += matrix[i, j];
                    worksheet.Cells[row1 + 1 + 2 + i, j + 2].Value = (matrix[i, j]);
                }
                worksheet.Cells[row1 + 1 + 2 + i, cols + 2].Value = rowstotal;
            }
            worksheet.Cells[row1 + 1 + 2 + rows, 1].Value = "Sum";


            for (int j = 0; j < cols; j++)
            {
                double colstotal = 0;

                for (int i = 0; i < rows; i++)
                {
                    colstotal += matrix[i, j];
                }
                worksheet.Cells[row1 + 1 + 2 + rows, j + 2].Value = colstotal;
            }
            worksheet.Cells[row1 + 1 + 2 + rows, cols + 2].Value = matrixTotal;
            CreateExcelWithColorScale(worksheet, row1 + 2, 1, row1 + 1 + 1 + rows, cols + 2);

            int row2 = row1 + 1 + 2 + rows;
            int col2 = cols + 2;

            // 设置第一行的字体为微软雅黑、大小为14号
            using (ExcelRange range1 = worksheet.Cells["A1:" + NumberToExcelColumn(col2) + $"{row2}"])
            {
                range1.Style.Font.Name = "微软雅黑";
                range1.Style.Font.Size = 11;
            }

            worksheet.Cells.AutoFitColumns();

            // 设置不显示网格线
            worksheet.View.ShowGridLines = false;

            // 确保文件名不为空
            if (!string.IsNullOrEmpty(output_excel_file))
            {
                FileInfo excelFile = new FileInfo(output_excel_file);

                // 检查文件是否已存在
                if (excelFile.Exists)
                {
                    try
                    {
                        // 文件未被打开，继续保存 Excel 文件
                        excelPackage.SaveAs(excelFile);
                    }
                    catch (IOException)
                    {
                        // Excel 文件已存在并且被打开，弹出提示框显示
                        MessageBox.Show($"文件 {output_excel_file} 已存在并且被打开，请关闭后重新保存。", "文件已打开", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                        EnableAllButtons();
                        return;
                    }
                }
                else
                {
                    // 文件不存在，直接保存 Excel 文件
                    excelPackage.SaveAs(excelFile);
                }
            }
        }

        String NumberToExcelColumn(int number)
        {
            string columnName = string.Empty;
            while (number > 0)
            {
                number--;
                columnName = (char)('A' + (number % 26)) + columnName;
                number /= 26;
            }
            return columnName;
        }

        void PrintAndWrite3DMatrix(int[,,] matrix, List<double[]>[] pairs, string output_excel_file)
        {
            string p1 = this.para1.Text.ToUpper();
            string p2 = this.para2.Text.ToUpper();
            string p3 = this.para3.Text.ToUpper();

            double total = 0;
            int rows = matrix.GetLength(0);
            int cols = matrix.GetLength(1);
            int depth = matrix.GetLength(2);

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage();
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("matrix3D");

            worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells[1, 1].Value = BinName.Text;
            worksheet.Cells[1, 2].Value = "片量";
            worksheet.Cells[1, 3].Value = waferList.Count;
            worksheet.Cells[1, 8].Value = "总计百分比";


            // 获取单元格范围对象
            var address = new ExcelAddress(1, 1, 1, 8);
            ExcelRange range = worksheet.Cells[address.Address];
            range.Style.Font.Bold = true;
            range.Style.Font.Color.SetColor(System.Drawing.Color.Red);

            for (int k = 0; k < depth; k++)
            {
                for (int j = 0; j < cols; j++)
                {

                    for (int i = 0; i < rows; i++)
                    {
                        total += matrix[i, j, k];
                    }
                }
            }

            int row11 = 1;

            for (int k = 0; k < depth; k++)
            {
                
                if (k == 0)
                {
                    worksheet.Cells[k * (rows +4) + 1 + row11, 1].Value = ($"{p3} <{pairs[2][0][0]}");
                    worksheet.Cells[k * (rows +4) + 2 + row11, 1].Value = ($"{p1}/{p2}");
                }
                else if (k == depth - 1)
                {
                    worksheet.Cells[k * (rows +4) + 1 + row11, 1].Value = ($"{p3} >{pairs[2][k-2][1]}");
                    worksheet.Cells[k * (rows +4) + 2 + row11, 1].Value = ($"{p1}/{p2}");
                }
                else
                {
                    worksheet.Cells[k * (rows +4) + 1 + row11, 1].Value = ($"{p3} {pairs[2][k - 1][0]}-{pairs[2][k - 1][1]}");
                    worksheet.Cells[k * (rows +4) + 2 + row11, 1].Value = ($"{p1}/{p2}");
                }

                int row1 = k * (rows + 4) + 1 + 1;

                for (int j = 0; j < cols; j++)
                {
                    if (j == 0)
                    {
                        worksheet.Cells[row1 + 1, j + 2].Value = ($"<{pairs[1][0][0]}");
                    }
                    else if (j == cols - 1)
                    {
                        worksheet.Cells[row1 + 1, j + 2].Value = ($">{pairs[1][j - 2][1]}");
                    }
                    else
                    {
                        worksheet.Cells[row1 + 1, j + 2].Value = ($"{pairs[1][j - 1][0]}-{pairs[1][j - 1][1]}");
                    }
                }
                worksheet.Cells[row1 + 1, cols + 2].Value = "Sum";
                for (int i = 0; i < rows; i++)
                {
                    double rowstotal = 0;

                    if (i == 0)
                    {
                        worksheet.Cells[row1 + 2 + i, 1].Value = ($"<{pairs[0][0][0]}");
                    }
                    else if (i == rows - 1)
                    {
                        worksheet.Cells[row1 + 2 + i, 1].Value = ($">{pairs[0][i - 2][1]}");
                    }
                    else
                    {
                        worksheet.Cells[row1 + 2 + i, 1].Value = ($"{pairs[0][i - 1][0]}-{pairs[0][i - 1][1]}");
                    }
                    for (int j = 0; j < cols; j++)
                    {
                        rowstotal += matrix[i, j, k];
                        worksheet.Cells[row1 + 2 + i, j + 2].Value = (matrix[i, j, k]) / total;
                        worksheet.Cells[row1 + 2 + i, j + 2].Style.Numberformat.Format = "0.00%";
                    }
                    worksheet.Cells[row1 + 2 + i, cols + 2].Value = rowstotal / total;
                    worksheet.Cells[row1 + 2 + i, cols + 2].Style.Numberformat.Format = "0.00%";

                }
                worksheet.Cells[row1 + 2 + rows, 1].Value = "Sum";
                CreateExcelWithColorScale(worksheet, row1 + 1, 1, row1 + 1 + rows, cols + 2);
                double matrixTotal = 0;
                for (int j = 0; j < cols; j++)
                {
                    double colstotal = 0; 
                    
                    for (int i = 0; i < rows; i++)
                    {
                        colstotal += matrix[i, j, k];
                        matrixTotal += matrix[i, j, k];
                    }
                    worksheet.Cells[row1 + 2 + rows, j + 2].Value = colstotal / total;
                    worksheet.Cells[row1 + 2 + rows, j + 2].Style.Numberformat.Format = "0.00%";
                }
                worksheet.Cells[row1 + 2 + rows, cols + 2].Value = matrixTotal / total;
                worksheet.Cells[row1 + 2 + rows, cols + 2].Style.Numberformat.Format = "0.00%";
            }

            int nextRows = (depth - 1) * (rows + 4) + 3 + rows + 2 + 1;

            for (int k = 0; k < depth; k++)
            {
                if (k == 0)
                {
                    worksheet.Cells[nextRows + k * (rows + 4) , 1].Value = BinName.Text;
                    worksheet.Cells[nextRows + k * (rows + 4) , 2].Value = "片量";
                    worksheet.Cells[nextRows + k * (rows + 4) , 3].Value = waferList.Count;

                    worksheet.Cells[nextRows + k * (rows + 4) , 8].Value = "颗粒数分布";
                    
                    // 获取单元格范围对象
                    address = new ExcelAddress(nextRows + k * (rows + 4), 1, nextRows + k * (rows + 4), 8);
                    range = worksheet.Cells[address.Address];
                    range.Style.Font.Bold = true;
                    range.Style.Font.Color.SetColor(System.Drawing.Color.Red);

                    worksheet.Cells[nextRows + k * (rows + 4) + 1,  1].Value = ($"{p3} <{pairs[2][0][0]}");
                    worksheet.Cells[nextRows + k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
                }
                else if (k == depth - 1)
                {
                    worksheet.Cells[nextRows + k * (rows + 4) + 1, 1].Value = ($"{p3} >{pairs[2][k - 2][1]}");
                    worksheet.Cells[nextRows + k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
                }
                else
                {
                    worksheet.Cells[nextRows + k * (rows + 4) + 1, 1].Value = ($"{p3} {pairs[2][k - 1][0]}-{pairs[2][k - 1][1]}");
                    worksheet.Cells[nextRows + k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
                }

                int row1 = k * (rows + 4) + 1;
                for (int j = 0; j < cols; j++)
                {
                    if (j == 0)
                    {
                        worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($"<{pairs[1][0][0]}");
                    }
                    else if (j == cols - 1)
                    {
                        worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($">{pairs[1][j - 2][1]}");
                    }
                    else
                    {
                        worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($"{pairs[1][j - 1][0]}-{pairs[1][j - 1][1]}");
                    }
                }
                worksheet.Cells[nextRows + row1 + 1, cols + 2].Value = "Sum";
                for (int i = 0; i < rows; i++)
                {
                    double rowstotal = 0;

                    if (i == 0)
                    {
                        worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($"<{pairs[0][0][0]}");
                    }
                    else if (i == rows - 1)
                    {
                        worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($">{pairs[0][i - 2][1]}");
                    }
                    else
                    {
                        worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($"{pairs[0][i - 1][0]}-{pairs[0][i - 1][1]}");
                    }
                    for (int j = 0; j < cols; j++)
                    {
                        rowstotal += matrix[i, j, k];
                        worksheet.Cells[nextRows + row1 + 2 + i, j + 2].Value = matrix[i, j, k];
                    }
                    worksheet.Cells[nextRows + row1 + 2 + i, cols + 2].Value = rowstotal;
                }
                worksheet.Cells[nextRows + row1 + 2 + rows, 1].Value = "Sum";
                CreateExcelWithColorScale(worksheet, nextRows + row1 + 1, 1, nextRows + row1 + 1 + rows, cols + 2);
                double matrixTotal = 0;
                for (int j = 0; j < cols; j++)
                {
                    double colstotal = 0;

                    for (int i = 0; i < rows; i++)
                    {
                        colstotal += matrix[i, j, k];
                        matrixTotal += matrix[i, j, k];
                    }
                    worksheet.Cells[nextRows + row1 + 2 + rows, j + 2].Value = colstotal;
                }
                worksheet.Cells[nextRows + row1 + 2 + rows, cols + 2].Value = matrixTotal;
            }

            int row2 = nextRows + (depth - 1)*(rows + 4) + 3 + rows;
            int col2 = cols + 2;

            using (ExcelRange range1 = worksheet.Cells["A1:"+ NumberToExcelColumn(col2) + $"{row2}"])
            {
                range1.Style.Font.Name = "微软雅黑";
                range1.Style.Font.Size = 11;
            }

            worksheet.Cells.AutoFitColumns();

            // 设置不显示网格线
            worksheet.View.ShowGridLines = false;

            if (!string.IsNullOrEmpty(output_excel_file))
            {
                FileInfo excelFile = new FileInfo(output_excel_file);

                if (excelFile.Exists)
                {
                    try
                    {
                        excelPackage.SaveAs(excelFile);
                    }
                    catch (IOException)
                    {
                        MessageBox.Show($"文件 {output_excel_file} 已存在并且被打开，请关闭后重新保存。", "文件已打开", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                        EnableAllButtons();
                        return;
                    }
                }
                else
                {
                    excelPackage.SaveAs(excelFile);
                }
            }
        }

        void PrintAndWrite4DMatrix(int[,,,] matrix, List<double[]>[] pairs, string output_excel_file)
        {

            //int rows = matrix.GetLength(0);
            //int cols = matrix.GetLength(1);
            //int depth = matrix.GetLength(2);
            //int time = matrix.GetLength(3);

            //string p1 = this.para1.Text;
            //string p2 = this.para2.Text;
            //string p3 = this.para3.Text;
            //string p4 = this.para4.Text;

            //double total = 0;

            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //ExcelPackage excelPackage = new ExcelPackage();
            //ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("matrix4D");

            //worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            //worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            //worksheet.Cells[1, 3].Value = ($"{waferList.Count}");

            //for (int t = 0; t < time; t++)
            //{
            //    if (t == 0)
            //    {
            //        worksheet.Cells[t * (rows + 4) + 1, 1].Value = ($"{p4} <{pairs[3][0][0]},");
            //        worksheet.Cells[t * (rows + 4) + 2, 1].Value = ($"{p3} <{pairs[2][0][0]}");
            //        worksheet.Cells[t * (rows + 4) + 3, 1].Value = ($"{p1}/{p2}");
            //    }
            //    else if (t == time - 1)
            //    {
            //        writer.WriteLine($"{p4} >{pairs[3][t - 2][1]},");
            //    }
            //    else
            //    {
            //        writer.WriteLine($"{p4} {pairs[3][t - 1][0]}-{pairs[3][t - 1][1]},");
            //    }
            //    for (int k = 0; k < depth; k++)
            //    {
            //        if (k == 0)
            //        {
            //            worksheet.Cells[k * (rows + 4) + 1, 1].Value = ($"{p3} <{pairs[2][0][0]}");
            //            worksheet.Cells[k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
            //        }
            //        else if (k == depth - 1)
            //        {
            //            worksheet.Cells[k * (rows + 4) + 1, 1].Value = ($"{p3} >{pairs[2][0][0]}");
            //            worksheet.Cells[k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
            //        }
            //        else
            //        {
            //            worksheet.Cells[k * (rows + 4) + 1, 1].Value = ($"{p3} {pairs[2][k - 1][0]}-{pairs[2][k - 1][1]}");
            //            worksheet.Cells[k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
            //        }
            //        int row1 = k * (rows + 4) + 1;
            //        for (int j = 0; j < cols; j++)
            //        {
            //            if (j == 0)
            //            {
            //                worksheet.Cells[row1 + 1, j + 2].Value = ($"<{pairs[1][0][0]}");
            //            }
            //            else if (j == cols - 1)
            //            {
            //                worksheet.Cells[row1 + 1, j + 2].Value = ($">{pairs[1][j - 2][1]}");
            //            }
            //            else
            //            {
            //                worksheet.Cells[row1 + 1, j + 2].Value = ($"{pairs[1][j - 1][0]}-{pairs[1][j - 1][1]}");
            //            }
            //        }
            //        worksheet.Cells[row1 + 1, cols + 2].Value = "Sum";
            //        for (int i = 0; i < rows; i++)
            //        {
            //            double rowstotal = 0;

            //            if (i == 0)
            //            {
            //                worksheet.Cells[row1 + 2 + i, 1].Value = ($"<{pairs[0][0][0]}");
            //            }
            //            else if (i == rows - 1)
            //            {
            //                worksheet.Cells[row1 + 2 + i, 1].Value = ($">{pairs[0][i - 2][1]}");
            //            }
            //            else
            //            {
            //                worksheet.Cells[row1 + 2 + i, 1].Value = ($"{pairs[0][i - 1][0]}-{pairs[0][i - 1][1]}");
            //            }
            //            for (int j = 0; j < cols; j++)
            //            {
            //                rowstotal += matrix[i, j, k];
            //                worksheet.Cells[row1 + 2 + i, j + 2].Value = (matrix[i, j, k]);
            //            }
            //            worksheet.Cells[row1 + 2 + i, cols + 2].Value = rowstotal;
            //        }
            //        worksheet.Cells[row1 + 2 + rows, 1].Value = "Sum";
            //        double matrixTotal = 0;
            //        for (int j = 0; j < cols; j++)
            //        {
            //            double colstotal = 0;

            //            for (int i = 0; i < rows; i++)
            //            {
            //                colstotal += matrix[i, j, k];
            //                matrixTotal += matrix[i, j, k];
            //            }
            //            worksheet.Cells[row1 + 2 + rows, j + 2].Value = colstotal;
            //        }
            //        worksheet.Cells[row1 + 2 + rows, cols + 2].Value = matrixTotal;
            //        total += matrixTotal;
            //    }
            //}

            //int nextRows = (depth - 1) * (rows + 4) + 3 + rows + 2;
            //for (int k = 0; k < depth; k++)
            //{
            //    if (k == 0)
            //    {
            //        worksheet.Cells[nextRows + k * (rows + 4) + 1, 1].Value = ($"{p3} <{pairs[2][0][0]}");
            //        worksheet.Cells[nextRows + k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
            //    }
            //    else if (k == depth - 1)
            //    {
            //        worksheet.Cells[nextRows + k * (rows + 4) + 1, 1].Value = ($"{p3} <{pairs[2][0][0]}");
            //        worksheet.Cells[nextRows + k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
            //    }
            //    else
            //    {
            //        worksheet.Cells[nextRows + k * (rows + 4) + 1, 1].Value = ($"{p3} {pairs[2][k - 1][0]}-{pairs[2][k - 1][1]}");
            //        worksheet.Cells[nextRows + k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
            //    }
            //    int row1 = k * (rows + 4) + 1;
            //    for (int j = 0; j < cols; j++)
            //    {
            //        if (j == 0)
            //        {
            //            worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($"<{pairs[1][0][0]}");
            //        }
            //        else if (j == cols - 1)
            //        {
            //            worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($">{pairs[1][j - 2][1]}");
            //        }
            //        else
            //        {
            //            worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($"{pairs[1][j - 1][0]}-{pairs[1][j - 1][1]}");
            //        }
            //    }
            //    worksheet.Cells[nextRows + row1 + 1, cols + 2].Value = "Sum";
            //    for (int i = 0; i < rows; i++)
            //    {
            //        double rowstotal = 0;

            //        if (i == 0)
            //        {
            //            worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($"<{pairs[0][0][0]}");
            //        }
            //        else if (i == rows - 1)
            //        {
            //            worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($">{pairs[0][i - 2][1]}");
            //        }
            //        else
            //        {
            //            worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($"{pairs[0][i - 1][0]}-{pairs[0][i - 1][1]}");
            //        }
            //        for (int j = 0; j < cols; j++)
            //        {
            //            rowstotal += matrix[i, j, k];
            //            worksheet.Cells[nextRows + row1 + 2 + i, j + 2].Value = matrix[i, j, k] / total;
            //            worksheet.Cells[nextRows + row1 + 2 + i, j + 2].Style.Numberformat.Format = "0.00%";
            //        }
            //        worksheet.Cells[nextRows + row1 + 2 + i, cols + 2].Value = rowstotal / total;
            //        worksheet.Cells[nextRows + row1 + 2 + i, cols + 2].Style.Numberformat.Format = "0.00%";
            //    }
            //    worksheet.Cells[nextRows + row1 + 2 + rows, 1].Value = "Sum";
            //    double matrixTotal = 0;
            //    for (int j = 0; j < cols; j++)
            //    {
            //        double colstotal = 0;

            //        for (int i = 0; i < rows; i++)
            //        {
            //            colstotal += matrix[i, j, k];
            //            matrixTotal += matrix[i, j, k];
            //        }
            //        worksheet.Cells[nextRows + row1 + 2 + rows, j + 2].Value = colstotal / total;
            //        worksheet.Cells[nextRows + row1 + 2 + rows, j + 2].Style.Numberformat.Format = "0.00%";
            //    }
            //    worksheet.Cells[nextRows + row1 + 2 + rows, cols + 2].Value = matrixTotal / total;
            //    worksheet.Cells[nextRows + row1 + 2 + rows, cols + 2].Style.Numberformat.Format = "0.00%";
            //}
            //int row2 = nextRows + (depth - 1) * (rows + 4) + 3 + rows;
            //int col2 = cols + 2;

            //using (ExcelRange range = worksheet.Cells["A1:" + NumberToExcelColumn(col2) + $"{row2}"])
            //{
            //    range.Style.Font.Name = "微软雅黑";
            //    range.Style.Font.Size = 11;
            //}

            //worksheet.Cells.AutoFitColumns();

            //if (!string.IsNullOrEmpty(output_excel_file))
            //{
            //    FileInfo excelFile = new FileInfo(output_excel_file);

            //    if (excelFile.Exists)
            //    {
            //        try
            //        {
            //            excelPackage.SaveAs(excelFile);
            //        }
            //        catch (IOException)
            //        {
            //            MessageBox.Show($"文件 {output_excel_file} 已存在并且被打开，请关闭后重新保存。", "文件已打开", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            //            return;
            //        }
            //    }
            //    else
            //    {
            //        excelPackage.SaveAs(excelFile);
            //    }
            //}

            //for (int t = 0; t < time; t++)
            //{
            //    if(t ==0)
            //    {
            //        writer.WriteLine($"{p4} <{pairs[3][0][0]},");
            //    }
            //    else if(t == time - 1)
            //    {
            //        writer.WriteLine($"{p4} >{pairs[3][t-2][1]},");
            //    }
            //    else
            //    {
            //        writer.WriteLine($"{p4} {pairs[3][t-1][0]}-{pairs[3][t-1][1]},");
            //    }

            //    for (int k = 0; k < depth; k++)
            //    {
            //        if (k == 0)
            //        {
            //            writer.WriteLine($"{p3} <{pairs[2][0][0]},");
            //        }
            //        else if (k == depth - 1)
            //        {
            //            writer.WriteLine($"{p3} >{pairs[2][k-2][1]},");
            //        }
            //        else
            //        {
            //            writer.WriteLine($"{p3} {pairs[2][k-1][0]}-{pairs[2][k-1][1]},");
            //        }


            //        writer.Write(" ");
            //        for (int j = 0; j < cols; j++)
            //        {
            //            if (j == 0)
            //            {
            //                writer.Write($"{p1}/{p2},<{pairs[1][j][1]},");
            //            }
            //            else if(j == cols - 1)
            //            {
            //                writer.Write($">{pairs[1][j-2][1]},");
            //            }
            //            else
            //            {
            //                writer.Write($"{pairs[1][j-1][0]}-{pairs[1][j-1][1]},");
            //            }

            //        }
            //        writer.WriteLine();

            //        for (int i = 0; i < rows; i++)
            //        {
            //            if (i == 0)
            //            {
            //                writer.Write($"<{pairs[0][i][1]},");
            //            }
            //            else if (i == rows-1)
            //            {
            //                writer.Write($">{pairs[0][i-2][1]},");
            //            }
            //            else
            //            {
            //                writer.Write($"{pairs[0][i-1][0]}-{pairs[0][i - 1][1]},");
            //            }
            //            for (int j = 0; j < cols; j++)
            //            {
            //                writer.Write(matrix[i, j, k, t].ToString()+",");
            //            }
            //            writer.WriteLine();
            //        }
            //    }
            //}
        }
        void clearButton_Click(object sender, RoutedEventArgs e)
        {
            // 清空所有TextBox
            ClearAllTextBoxes(this);
        }

         void ClearAllTextBoxes(DependencyObject parent)
        {
            // 递归遍历所有子控件，找到所有的TextBox并清空它们的文本
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is System.Windows.Controls.TextBox textBox)
                {
                    textBox.Clear();
                }
                else
                {
                    ClearAllTextBoxes(child); // 递归调用
                }
            }
        }

        List<List<double[]>> GeneratePermutations(List<double[]>[] pairs)
        {
            List<List<double[]>> result = new List<List<double[]>>();

            // 使用递归生成全排列
            GeneratePermutationsRecursive(pairs, 0, new List<double[]>(), result);

            return result;
        }

        void GeneratePermutationsRecursive(List<double[]>[] pairs, int index, List<double[]> current, List<List<double[]>> result)
        {
            if (index == pairs.Length)
            {
                result.Add(new List<double[]>(current));
                return;
            }

            foreach (var pair in pairs[index])
            {
                current.Add(pair);
                GeneratePermutationsRecursive(pairs, index + 1, current, result);
                current.RemoveAt(current.Count - 1);
            }
        }

        void SaveToCsv(List<List<double[]>> data, StreamWriter writer, int[] col)
        {
            int rowIndex = 1;  // 用于第三列的排序编号

            foreach (var row in data)
            {
                // 创建包含24列的数组
                string[] line = new string[78];

                // 设置第一列为1
                line[0] = "1";

                // 设置第三列为排序编号
                line[2] = rowIndex.ToString();
                rowIndex++;

                // 按指定列填充数据
                if (row.Count > 0)
                {
                    line[col[0]] = row[0][0].ToString();
                    line[col[0]+1] = row[0][1].ToString();
                }
                if (row.Count > 1)
                {
                    line[col[1]] = row[1][0].ToString();
                    line[col[1]+1] = row[1][1].ToString();
                }
                if (row.Count > 2)
                {
                    line[col[2]] = row[2][0].ToString();
                    line[col[2]+1] = row[2][1].ToString();
                }
                if (row.Count > 3)
                {
                    line[col[3]] = row[3][0].ToString();
                    line[col[3]+1] = row[3][1].ToString();
                }

                // 将行写入CSV文件
                writer.WriteLine(string.Join(",", line));
            }
        }

        void genButton_Click(object sender, RoutedEventArgs e)
        {
            int dim = 0;
            if (!int.TryParse(dimensionTextBox.Text, out dim))
            {
                System.Windows.MessageBox.Show("Invalid dimension value. Please enter a valid integer.");
                EnableAllButtons();
                return;
            }

            string outputFolder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutputFolder");
            string JsonfilePath = System.IO.Path.Combine(outputFolder, "pairsConfig.json"); 
            
            // 检查文件夹是否存在
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }
            // 定义字典
            Dictionary<string, int> dict = new Dictionary<string, int>
            {
                {"VF1", 4}, {"VF2", 6}, {"VF3", 8}, {"VF4", 10}, {"VZ1", 12},
                {"IR", 14}, {"HW1", 16}, {"LOP1", 18}, {"WLP1", 20}, {"WLD1", 22},
                {"IR1", 24}, {"VFD", 26}, {"DVF", 28}, {"IR2", 30}, {"WLC1", 32},
                {"VF5", 34}, {"VF6", 36}, {"VF7", 38}, {"VF8", 40}, {"DVF1", 42},
                {"DVF2", 44}, {"VZ2", 46}, {"VZ3", 48}, {"VZ4", 50}, {"VZ5", 52},
                {"IR3", 54}, {"IR4", 56}, {"IR5", 58}, {"IR6", 60}, {"IF", 62},
                {"IF1", 64}, {"IF2", 66}, {"LOP2", 68}, {"WLP2", 70}, {"WLD2", 72},
                {"HW2", 74}, {"WLC2", 76}
            };

            // 获取文件名
            string fileName = BinName.Text + ".csv";
            fileName = System.IO.Path.Combine(outputFolder, fileName);

            // 写入文件内容
            using (StreamWriter writer = new StreamWriter(fileName))
            {
                // 写入每行内容
                writer.WriteLine("ResortBin");
                writer.WriteLine("Format1");
                writer.WriteLine("Bin Table Name,"+ BinName.Text);
                writer.WriteLine("");
                writer.WriteLine("BinSetting,"+ BinSetting.Text);
                writer.WriteLine("TestItems,,#,,VF1,VF1,VF2,VF2,VF3,VF3,VF4,VF4,VZ1,VZ1,IR,IR,HW1,HW1,LOP1,LOP1,WLP1,WLP1,WLD1,WLD1,IR1,IR1,VFD,VFD,DVF,DVF,IR2,IR2,WLC1,WLC1,VF5,VF5,VF6,VF6,VF7,VF7,VF8,VF8,DVF1,DVF1,DVF2,DVF2,VZ2,VZ2,VZ3,VZ3,VZ4,VZ4,VZ5,VZ5,IR3,IR3,IR4,IR4,IR5,IR5,IR6,IR6,IF,IF,IF1,IF1,IF2,IF2,LOP2,LOP2,WLP2,WLP2,WLD2,WLD2,HW2,HW2,WLC2,WLC2,");
                writer.WriteLine("BinEnable,BinType,BIN,Code,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,Min,Max,");
                try
                {
                    // 获取参数
                    int dimension = int.Parse(dimensionTextBox.Text);
                    if (dimension < 1 || dimension > 4)
                    {
                        MessageBox.Show("Dimension must be between 1 and 4.");
                        EnableAllButtons();
                        return;
                    }

                    double[] minValues = new double[dimension];
                    double[] stepSizes = new double[dimension];
                    int[] counts = new int[dimension];
                    int[] col = new int[dimension];
                    string[] paraStrings = new string[dimension];
                    if (dimension >= 1)
                    {
                        if (dict.ContainsKey(this.para1.Text.ToUpper())) col[0] = dict[this.para1.Text.ToUpper()];
                        paraStrings[0] = this.para1.Text.ToUpper();
                        minValues[0] = double.Parse(this.para1min.Text);
                        stepSizes[0] = double.Parse(this.para1rta.Text);
                        counts[0] = int.Parse(this.para1num.Text);
                    }
                    if (dimension >= 2)
                    {
                        if (dict.ContainsKey(this.para2.Text.ToUpper())) col[1] = dict[this.para2.Text.ToUpper()];
                        paraStrings[1] = this.para2.Text.ToUpper();
                        minValues[1] = double.Parse(this.para2min.Text);
                        stepSizes[1] = double.Parse(this.para2rta.Text);
                        counts[1] = int.Parse(this.para2num.Text);
                    }
                    if (dimension >= 3)
                    {
                        if (dict.ContainsKey(this.para3.Text.ToUpper())) col[2] = dict[this.para3.Text.ToUpper()];
                        paraStrings[2] = this.para3.Text.ToUpper();
                        minValues[2] = double.Parse(this.para3min.Text);
                        stepSizes[2] = double.Parse(this.para3rta.Text);
                        counts[2] = int.Parse(this.para3num.Text);
                    }
                    if (dimension >= 4)
                    {
                        if (dict.ContainsKey(this.para4.Text.ToUpper())) col[3] = dict[this.para4.Text.ToUpper()];
                        paraStrings[3] = this.para4.Text.ToUpper();
                        minValues[3] = double.Parse(this.para4min.Text);
                        stepSizes[3] = double.Parse(this.para4rta.Text);
                        counts[3] = int.Parse(this.para4num.Text);
                    }

                    var pairsDict = new Dictionary<string, List<double[]>>();

                    // 存储生成的pair数组
                    List<double[]>[] pairs = new List<double[]>[dimension];

                    // 生成pair数组
                    for (int i = 0; i < dimension; i++)
                    {
                        pairs[i] = new List<double[]>();
                        for (int j = 0; j < counts[i]; j++)
                        {
                            double first = Math.Round(minValues[i] + j * stepSizes[i], 6);
                            double second = Math.Round(first + stepSizes[i], 6);
                            pairs[i].Add(new double[] { first, second });
                        }
                        pairsDict[paraStrings[i]] = pairs[i];
                    }

                    ExportToCustomJson(pairsDict, JsonfilePath);

                    // 生成全排列结果
                    var result = GeneratePermutations(pairs);
                            SaveToCsv(result, writer,col);
                        }
                        catch (Exception ex)
                        {
                            System.Windows.MessageBox.Show($"生成数组时出现错误：{ex.Message}");
                        }
                    }

            // 弹出消息框询问是否打开文件
            MessageBoxResult Getresult = System.Windows.MessageBox.Show("Excel 文件已导出到 " + fileName + "\n是否打开该文件？", "导出成功", MessageBoxButton.YesNo, MessageBoxImage.Question);

            // 根据用户的选择执行相应的操作
            if (Getresult == MessageBoxResult.Yes)
            {
                // 打开文件
                try
                {
                    Process.Start(new ProcessStartInfo { FileName = fileName, UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("打开文件失败：" + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}
