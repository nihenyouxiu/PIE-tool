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


namespace 产出分布计算
{
    /// <summary>
    /// Page1.xaml 的交互逻辑
    /// </summary>
    public partial class Page1 : Page
    {
        private readonly object parameterlockObject = new object();
        readonly object lockObject = new object();
        class Chip
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

        public Page1()
        {
            InitializeComponent();
            this.KeepAlive = true;
        }

        void WriteMatrix(List<(double, double)>[] pairs, int dim, string output_excel_file)
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

        private async void ProcessFile(string filename, int dim, List<(double, double)>[] pairs, int[] col2)
        {
            List<Chip> chipList = new List<Chip>();
            int waferidchipnum = 0;
            int flag = 0;
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
                            waferidchipnum++;
                            if (dim >= 1)
                            {
                                Dimension1 = !string.IsNullOrEmpty(values[col2[0]+ flag]) ? Convert.ToDouble(values[col2[0] + flag]) : -100000;
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
                            chipList.Add(chipData);
                        }
                    }
                }
            }
            catch (IOException)
            {
                //await Dispatcher.InvokeAsync(() =>
                //{
                    MessageBox.Show($"文件 {filename} 已被打开，请关闭后重新选择!", "文件已打开", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                //});
                return;
            }

            if (chipList.Any())
            {

                if (dim == 1)
                {
                    // 处理一维数据
                    lock (lockObject)
                    {
                        Generate1DMatrix(chipList, pairs);
                    }
                }
                else if (dim == 2)
                {
                    lock (lockObject) 
                    { 
                        // 处理二维数据
                        Generate2DMatrix(chipList, pairs);
                    }
                }
                else if (dim == 3)
                {
                    lock (lockObject)
                    { 
                        // 处理三维数据
                        Generate3DMatrix(chipList, pairs);
                    }
                }
                else if (dim == 4)
                {
                    lock (lockObject)
                    {
                        // 处理四维数据
                        Generate4DMatrix(chipList, pairs);
                    }
                }
                chipList.Clear();
                await Dispatcher.InvokeAsync(() =>
                {
                    lock (parameterlockObject)
                    {
                        parameterListBox.Items.Add(Path.GetFileName(filename) + " 计算完成!");
                        // 滚动到最新项
                        parameterListBox.ScrollIntoView(parameterListBox.Items[parameterListBox.Items.Count - 1]);
                    }
                });
            }
            else
            {
                breakFlag = true;
                MessageBox.Show("输入文件有误，请重新输入！");
            }

        }

        int[] matrix1D;
        int[,] matrix2D;
        int[,,] matrix3D;
        int[,,,] matrix4D;
        async void runButton_Click(object sender, RoutedEventArgs e)
        {
            int dim = 0;
            string filePath = BinName.Text;
            if (!int.TryParse(dimensionTextBox.Text, out dim))
            {
                System.Windows.MessageBox.Show("Invalid dimension value. Please enter a valid integer.");
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
                throw new ArgumentException("Dimension must be between 1 and 4.");
            }

            double[] minValues = new double[dimension];
            double[] stepSizes = new double[dimension];
            int[] counts = new int[dimension];
            int[] col = new int[dimension];
            int[] col2 = new int[dimension];

            if (dimension >= 1)
            {
                if (dict.ContainsKey(this.para1.Text.ToUpper())) col[0] = dict[this.para1.Text.ToUpper()];
                if (dict2.ContainsKey(this.para1.Text.ToUpper())) col2[0] = dict2[this.para1.Text.ToUpper()];
                minValues[0] = double.Parse(this.para1min.Text);
                stepSizes[0] = double.Parse(this.para1rta.Text);
                counts[0] = int.Parse(this.para1num.Text);
            }
            if (dimension >= 2)
            {
                if (dict.ContainsKey(this.para2.Text.ToUpper())) col[1] = dict[this.para2.Text.ToUpper()]; 
                if (dict2.ContainsKey(this.para2.Text.ToUpper())) col2[1] = dict2[this.para2.Text.ToUpper()];
                minValues[1] = double.Parse(this.para2min.Text);
                stepSizes[1] = double.Parse(this.para2rta.Text);
                counts[1] = int.Parse(this.para2num.Text);
            }
            if (dimension >= 3)
            {
                if (dict.ContainsKey(this.para3.Text.ToUpper())) col[2] = dict[this.para3.Text.ToUpper()];
                if (dict2.ContainsKey(this.para3.Text.ToUpper())) col2[2] = dict2[this.para3.Text.ToUpper()];
                minValues[2] = double.Parse(this.para3min.Text);
                stepSizes[2] = double.Parse(this.para3rta.Text);
                counts[2] = int.Parse(this.para3num.Text);
            }
            if (dimension >= 4)
            {
                MessageBox.Show("4维数据产出分布未开发！");
                return;
                if (dict.ContainsKey(this.para4.Text.ToUpper())) col[3] = dict[this.para4.Text.ToUpper()];
                if (dict2.ContainsKey(this.para4.Text.ToUpper())) col2[3] = dict2[this.para4.Text.ToUpper()];
                minValues[3] = double.Parse(this.para4min.Text);
                stepSizes[3] = double.Parse(this.para4rta.Text);
                counts[3] = int.Parse(this.para4num.Text);
            }

            // 存储生成的pair数组
            List<(double, double)>[] pairs = new List<(double, double)>[dimension];

            // 生成pair数组
            for (int i = 0; i < dimension; i++)
            {
                pairs[i] = new List<(double, double)>();
                for (int j = 0; j < counts[i]; j++)
                {
                    double first = Math.Round(minValues[i] + j * stepSizes[i],3);
                    double second = Math.Round(first + stepSizes[i], 3);
                    pairs[i].Add((first, second));
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

            // 处理并输出结果
            
            string outputFolder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutputFolder");
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            string output_excel_file = System.IO.Path.Combine(outputFolder, $"{filePath}.xlsx");
            // 检查文件夹是否存在
            if (Directory.Exists(outputFolder))
            {
                // 尝试获取文件夹中的所有文件
                string[] files = Directory.GetFiles(outputFolder);

                // 遍历文件夹中的所有文件
                foreach (string file in files)
                {
                    try
                    {
                        // 尝试打开文件，如果文件已经被打开会引发 IOException 异常
                        using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.None))
                        {
                            // 文件未被打开，继续处理
                        }
                    }
                    catch (IOException)
                    {
                        // 文件被打开，弹出提示框显示
                        MessageBox.Show($"文件 {file} 已被打开，请关闭后重新尝试删除文件夹。", "文件已打开", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                        return; // 终止方法的执行，不继续删除文件夹
                    }
                }

                // 删除文件夹及其内容
                Directory.Delete(outputFolder, true);
            }

            // 创建文件夹
            Directory.CreateDirectory(outputFolder);
            if (openFileDialog.ShowDialog() == true)
            {
                DateTime startTime = DateTime.Now; // 记录开始时间

                List<Task> tasks = new List<Task>(); // 声明 tasks 列表

                // 尝试打开文件，如果文件已经被打开会引发 IOException 异常
                foreach (string filename in openFileDialog.FileNames)
                {
                    tasks.Add( Task.Run(() => ProcessFile(filename, dimension, pairs,col2))); // 使用多线程处理文件
                    if (breakFlag)
                    {
                        break;
                    }
                }
                await Task.WhenAll(tasks); // 等待所有任务完成
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

            }
            else
            {
                MessageBox.Show("请输入文件！");
            }
        }
        void Generate1DMatrix(List<Chip> chips, List<(double, double)>[] pairs)
        {

            foreach (var chip in chips)
            {
                int index = GetRangeIndex(chip.GetDimensionValue(0), pairs[0]);
                if (index >= 0)
                {
                    matrix1D[index]++;
                }
            }
        }

        void Generate2DMatrix(List<Chip> chips, List<(double, double)>[] pairs)
        {

            foreach (var chip in chips)
            {
                int rowIndex = GetRangeIndex(chip.GetDimensionValue(0), pairs[0]);
                int colIndex = GetRangeIndex(chip.GetDimensionValue(1), pairs[1]);

                if (rowIndex >= 0 && colIndex >= 0)
                {
                    matrix2D[rowIndex, colIndex]++;
                }
            }
        }

        void Generate3DMatrix(List<Chip> chips, List<(double, double)>[] pairs)
        {
            foreach (var chip in chips)
            {
                int rowIndex = GetRangeIndex(chip.GetDimensionValue(0), pairs[0]);
                int colIndex = GetRangeIndex(chip.GetDimensionValue(1), pairs[1]);
                int depthIndex = GetRangeIndex(chip.GetDimensionValue(2), pairs[2]);

                if (rowIndex >= 0 && colIndex >= 0 && depthIndex >= 0)
                {
                    matrix3D[rowIndex, colIndex, depthIndex]++;
                }
            }
        }

        void Generate4DMatrix(List<Chip> chips, List<(double, double)>[] pairs)
        {

            foreach (var chip in chips)
            {
                int rowIndex = GetRangeIndex(chip.GetDimensionValue(0), pairs[0]);
                int colIndex = GetRangeIndex(chip.GetDimensionValue(1), pairs[1]);
                int depthIndex = GetRangeIndex(chip.GetDimensionValue(2), pairs[2]);
                int timeIndex = GetRangeIndex(chip.GetDimensionValue(3), pairs[3]);

                if (rowIndex >= 0 && colIndex >= 0 && depthIndex >= 0 && timeIndex >= 0)
                {
                    matrix4D[rowIndex, colIndex, depthIndex, timeIndex]++;
                }
            }
        }

        int GetRangeIndex(double value, List<(double, double)> pairs)
        {
            for (int i = 0; i < pairs.Count; i++)
            {
                if (value >= pairs[i].Item1 && value < pairs[i].Item2)
                {
                    return i + 1;
                }else if(value < pairs[0].Item1)
                {
                    return 0;
                }else if(value >= pairs[pairs.Count - 1].Item2)
                {
                    return pairs.Count+1;
                }
            }
            return -1;  // 不在任何范围内
        }

        void PrintAndWrite1DMatrix(int[] matrix, List<(double, double)>[] pairs, string output_excel_file)
        {
            string p1 = this.para1.Text.ToUpper();

            int rows = matrix.GetLength(0);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage();
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("matrix1D");

            worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells[1, 1].Value = ($"{p1}");
            int row1 = 0;
            double rowstotal = 0;
            
            for (int i = 0; i < rows; i++)
            {

                if (i == 0)
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($"<{pairs[0][0].Item1}");
                }
                else if (i == rows - 1)
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($">{pairs[0][i - 2].Item2}");
                }
                else
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($"{pairs[0][i - 1].Item1}-{pairs[0][i - 1].Item2}");
                }

                rowstotal += matrix[i];
                worksheet.Cells[row1 + 2 + i, 2].Value = (matrix[i]);
            }
            worksheet.Cells[row1 + 2 + rows, 1].Value = "Sum";
            worksheet.Cells[row1 + 2 + rows, 2].Value = rowstotal;

            int row2 = row1 + 2 + rows;
            int col2 = 2;
            // 设置第一行的字体为微软雅黑、大小为14号
            using (ExcelRange range = worksheet.Cells["A1:" + NumberToExcelColumn(col2) + $"{row2}"])
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
            ExcelRange range = worksheet.Cells[address.Address];

            // 设置边框样式
            range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        }

        void PrintAndWrite2DMatrix(int[,] matrix, List<(double, double)>[] pairs, string output_excel_file)
        {
            string p1 = this.para1.Text.ToUpper();
            string p2 = this.para2.Text.ToUpper();

            double total = 0;
            int rows = matrix.GetLength(0);
            int cols = matrix.GetLength(1);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage();
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("matrix2D");

            worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells[1, 1].Value = ($"{p1}/{p2}");
            int row1 = 0;
            for (int j = 0; j < cols; j++)
            {
                if (j == 0)
                {
                    worksheet.Cells[row1 + 1, j + 2].Value = ($"<{pairs[1][0].Item1}");
                }
                else if (j == cols - 1)
                {
                    worksheet.Cells[row1 + 1, j + 2].Value = ($">{pairs[1][j - 2].Item2}");
                }
                else
                {
                    worksheet.Cells[row1 + 1, j + 2].Value = ($"{pairs[1][j - 1].Item1}-{pairs[1][j - 1].Item2}");
                }
            }
            worksheet.Cells[row1 + 1, cols + 2].Value = "Sum";
            for (int i = 0; i < rows; i++)
            {
                double rowstotal = 0;

                if (i == 0)
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($"<{pairs[0][0].Item1}");
                }
                else if (i == rows - 1)
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($">{pairs[0][i - 2].Item2}");
                }
                else
                {
                    worksheet.Cells[row1 + 2 + i, 1].Value = ($"{pairs[0][i - 1].Item1}-{pairs[0][i - 1].Item2}");
                }
                for (int j = 0; j < cols; j++)
                {
                    rowstotal += matrix[i, j];
                    worksheet.Cells[row1 + 2 + i, j + 2].Value = (matrix[i, j]);
                }
                worksheet.Cells[row1 + 2 + i, cols + 2].Value = rowstotal;
            }
            worksheet.Cells[row1 + 2 + rows, 1].Value = "Sum";
            CreateExcelWithColorScale(worksheet, 1, 1, row1 + 2 + rows, cols + 2);
            double matrixTotal = 0;
            for (int j = 0; j < cols; j++)
            {
                double colstotal = 0;

                for (int i = 0; i < rows; i++)
                {
                    colstotal += matrix[i, j];
                    matrixTotal += matrix[i, j];
                }
                worksheet.Cells[row1 + 2 + rows, j + 2].Value = colstotal;
            }
            worksheet.Cells[row1 + 2 + rows, cols + 2].Value = matrixTotal;
            total += matrixTotal;

            int nextRows = rows + 3;
            worksheet.Cells[nextRows + 1, 1].Value = ($"{p1}/{p2}");
            for (int j = 0; j < cols; j++)
            {
                if (j == 0)
                {
                    worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($"<{pairs[1][0].Item1}");
                }
                else if (j == cols - 1)
                {
                    worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($">{pairs[1][j - 2].Item2}");
                }
                else
                {
                    worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($"{pairs[1][j - 1].Item1}-{pairs[1][j - 1].Item2}");
                }
            }
            worksheet.Cells[nextRows + row1 + 1, cols + 2].Value = "Sum";
            for (int i = 0; i < rows; i++)
            {
                double rowstotal = 0;

                if (i == 0)
                {
                    worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($"<{pairs[0][0].Item1}");
                }
                else if (i == rows - 1)
                {
                    worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($">{pairs[0][i - 2].Item2}");
                }
                else
                {
                    worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($"{pairs[0][i - 1].Item1}-{pairs[0][i - 1].Item2}");
                }
                for (int j = 0; j < cols; j++)
                {
                    rowstotal += matrix[i, j];
                    worksheet.Cells[nextRows + row1 + 2 + i, j + 2].Value = (matrix[i, j]) / total;
                    worksheet.Cells[nextRows + row1 + 2 + i, j + 2].Style.Numberformat.Format = "0.00%";
                }
                worksheet.Cells[nextRows + row1 + 2 + i, cols + 2].Value = rowstotal / total;
                worksheet.Cells[nextRows + row1 + 2 + i, cols + 2].Style.Numberformat.Format = "0.00%";
            }
            worksheet.Cells[nextRows + row1 + 2 + rows, 1].Value = "Sum";
            CreateExcelWithColorScale(worksheet, nextRows + 1, 1, nextRows + row1 + 2 + rows, cols + 2);

            for (int j = 0; j < cols; j++)
            {
                double colstotal = 0;

                for (int i = 0; i < rows; i++)
                {
                    colstotal += matrix[i, j];
                }
                worksheet.Cells[nextRows + row1 + 2 + rows, j + 2].Value = colstotal / total;
                worksheet.Cells[nextRows + row1 + 2 + rows, j + 2].Style.Numberformat.Format = "0.00%";
            }
            worksheet.Cells[nextRows + row1 + 2 + rows, cols + 2].Value = matrixTotal / total;
            worksheet.Cells[nextRows + row1 + 2 + rows, cols + 2].Style.Numberformat.Format = "0.00%";


            int row2 = nextRows + row1 + 2 + rows;
            int col2 = cols + 2;
            // 设置第一行的字体为微软雅黑、大小为14号
            using (ExcelRange range = worksheet.Cells["A1:" + NumberToExcelColumn(col2) + $"{row2}"])
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

        void PrintAndWrite3DMatrix(int[,,] matrix, List<(double, double)>[] pairs, string output_excel_file)
        {
            string p1 = this.para1.Text.ToUpper();
            string p2 = this.para2.Text.ToUpper();
            string p3 = this.para3.Text.ToUpper();

            double total = 0;
            int rows = matrix.GetLength(0);
            int cols = matrix.GetLength(1);
            int depth = matrix.GetLength(2);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage();
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("matrix3D");

            worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            for (int k = 0; k < depth; k++)
            {
                if (k == 0)
                {
                    worksheet.Cells[k * (rows +4) + 1,1].Value = ($"{p3} <{pairs[2][0].Item1}");
                    worksheet.Cells[k * (rows +4) + 2,1].Value = ($"{p1}/{p2}");
                }
                else if (k == depth - 1)
                {
                    worksheet.Cells[k * (rows +4) + 1,1].Value = ($"{p3} >{pairs[2][k-2].Item2}");
                    worksheet.Cells[k * (rows +4) + 2,1].Value = ($"{p1}/{p2}");
                }
                else
                {
                    worksheet.Cells[k * (rows +4) + 1,1].Value = ($"{p3} {pairs[2][k - 1].Item1}-{pairs[2][k - 1].Item2}");
                    worksheet.Cells[k * (rows +4) + 2,1].Value = ($"{p1}/{p2}");
                }
                int row1 = k * (rows + 4) + 1;
                for (int j = 0; j < cols; j++)
                {
                    if (j == 0)
                    {
                        worksheet.Cells[row1 + 1, j + 2].Value = ($"<{pairs[1][0].Item1}");
                    }
                    else if (j == cols - 1)
                    {
                        worksheet.Cells[row1 + 1, j + 2].Value = ($">{pairs[1][j - 2].Item2}");
                    }
                    else
                    {
                        worksheet.Cells[row1 + 1, j + 2].Value = ($"{pairs[1][j - 1].Item1}-{pairs[1][j - 1].Item2}");
                    }
                }
                worksheet.Cells[row1 + 1, cols + 2].Value = "Sum";
                for (int i = 0; i < rows; i++)
                {
                    double rowstotal = 0;

                    if (i == 0)
                    {
                        worksheet.Cells[row1 + 2 + i, 1].Value = ($"<{pairs[0][0].Item1}");
                    }
                    else if (i == rows - 1)
                    {
                        worksheet.Cells[row1 + 2 + i, 1].Value = ($">{pairs[0][i - 2].Item2}");
                    }
                    else
                    {
                        worksheet.Cells[row1 + 2 + i, 1].Value = ($"{pairs[0][i - 1].Item1}-{pairs[0][i - 1].Item2}");
                    }
                    for (int j = 0; j < cols; j++)
                    {
                        rowstotal += matrix[i, j, k];
                        worksheet.Cells[row1 + 2 + i, j + 2].Value = (matrix[i, j, k]);
                    }
                    worksheet.Cells[row1 + 2 + i, cols + 2].Value = rowstotal;
                }
                worksheet.Cells[row1 + 2 + rows, 1].Value = "Sum";
                CreateExcelWithColorScale(worksheet, row1 + 1, 1, row1 + 2 + rows, cols + 2);
                double matrixTotal = 0;
                for (int j = 0; j < cols; j++)
                {
                    double colstotal = 0; 
                    
                    for (int i = 0; i < rows; i++)
                    {
                        colstotal += matrix[i, j, k];
                        matrixTotal += matrix[i, j, k];
                    }
                    worksheet.Cells[row1 + 2 + rows, j+2].Value = colstotal;
                }
                worksheet.Cells[row1 + 2 + rows, cols + 2].Value = matrixTotal;
                total += matrixTotal;
            }

            int nextRows = (depth - 1) * (rows + 4) + 3 + rows + 2;

            for (int k = 0; k < depth; k++)
            {
                if (k == 0)
                {
                    worksheet.Cells[nextRows + k * (rows + 4) + 1,  1].Value = ($"{p3} <{pairs[2][0].Item1}");
                    worksheet.Cells[nextRows + k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
                }
                else if (k == depth - 1)
                {
                    worksheet.Cells[nextRows + k * (rows + 4) + 1, 1].Value = ($"{p3} >{pairs[2][k - 2].Item2}");
                    worksheet.Cells[nextRows + k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
                }
                else
                {
                    worksheet.Cells[nextRows + k * (rows + 4) + 1, 1].Value = ($"{p3} {pairs[2][k - 1].Item1}-{pairs[2][k - 1].Item2}");
                    worksheet.Cells[nextRows + k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
                }
                int row1 = k * (rows + 4) + 1;
                for (int j = 0; j < cols; j++)
                {
                    if (j == 0)
                    {
                        worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($"<{pairs[1][0].Item1}");
                    }
                    else if (j == cols - 1)
                    {
                        worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($">{pairs[1][j - 2].Item2}");
                    }
                    else
                    {
                        worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($"{pairs[1][j - 1].Item1}-{pairs[1][j - 1].Item2}");
                    }
                }
                worksheet.Cells[nextRows + row1 + 1, cols + 2].Value = "Sum";
                for (int i = 0; i < rows; i++)
                {
                    double rowstotal = 0;

                    if (i == 0)
                    {
                        worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($"<{pairs[0][0].Item1}");
                    }
                    else if (i == rows - 1)
                    {
                        worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($">{pairs[0][i - 2].Item2}");
                    }
                    else
                    {
                        worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($"{pairs[0][i - 1].Item1}-{pairs[0][i - 1].Item2}");
                    }
                    for (int j = 0; j < cols; j++)
                    {
                        rowstotal += matrix[i, j, k];
                        worksheet.Cells[nextRows + row1 + 2 + i, j + 2].Value = matrix[i, j, k] / total;
                        worksheet.Cells[nextRows + row1 + 2 + i, j + 2].Style.Numberformat.Format = "0.00%";
                    }
                    worksheet.Cells[nextRows + row1 + 2 + i, cols + 2].Value = rowstotal / total;
                    worksheet.Cells[nextRows + row1 + 2 + i, cols + 2].Style.Numberformat.Format = "0.00%";
                }
                worksheet.Cells[nextRows + row1 + 2 + rows, 1].Value = "Sum";
                CreateExcelWithColorScale(worksheet, nextRows + row1 + 1, 1, nextRows + row1 + 2 + rows, cols + 2);
                double matrixTotal = 0;
                for (int j = 0; j < cols; j++)
                {
                    double colstotal = 0;

                    for (int i = 0; i < rows; i++)
                    {
                        colstotal += matrix[i, j, k];
                        matrixTotal += matrix[i, j, k];
                    }
                    worksheet.Cells[nextRows + row1 + 2 + rows, j + 2].Value = colstotal / total;
                    worksheet.Cells[nextRows + row1 + 2 + rows, j + 2].Style.Numberformat.Format = "0.00%";
                }
                worksheet.Cells[nextRows + row1 + 2 + rows, cols + 2].Value = matrixTotal / total;
                worksheet.Cells[nextRows + row1 + 2 + rows, cols + 2].Style.Numberformat.Format = "0.00%";
            }
            int row2 = nextRows + (depth - 1)*(rows + 4) + 3 + rows;
            int col2 = cols + 2;

            using (ExcelRange range = worksheet.Cells["A1:"+ NumberToExcelColumn(col2) + $"{row2}"])
            {
                range.Style.Font.Name = "微软雅黑";
                range.Style.Font.Size = 11;
            }

            worksheet.Cells.AutoFitColumns();

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
                        return;
                    }
                }
                else
                {
                    excelPackage.SaveAs(excelFile);
                }
            }
        }

        void PrintAndWrite4DMatrix(int[,,,] matrix, List<(double, double)>[] pairs, string output_excel_file)
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

            //for (int t = 0; t < time; t++)
            //{
            //    if (t == 0)
            //    {
            //        worksheet.Cells[t * (rows + 4) + 1, 1].Value = ($"{p4} <{pairs[3][0].Item1},");
            //        worksheet.Cells[t * (rows + 4) + 2, 1].Value = ($"{p3} <{pairs[2][0].Item1}");
            //        worksheet.Cells[t * (rows + 4) + 3, 1].Value = ($"{p1}/{p2}");
            //    }
            //    else if (t == time - 1)
            //    {
            //        writer.WriteLine($"{p4} >{pairs[3][t - 2].Item2},");
            //    }
            //    else
            //    {
            //        writer.WriteLine($"{p4} {pairs[3][t - 1].Item1}-{pairs[3][t - 1].Item2},");
            //    }
            //    for (int k = 0; k < depth; k++)
            //    {
            //        if (k == 0)
            //        {
            //            worksheet.Cells[k * (rows + 4) + 1, 1].Value = ($"{p3} <{pairs[2][0].Item1}");
            //            worksheet.Cells[k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
            //        }
            //        else if (k == depth - 1)
            //        {
            //            worksheet.Cells[k * (rows + 4) + 1, 1].Value = ($"{p3} >{pairs[2][0].Item1}");
            //            worksheet.Cells[k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
            //        }
            //        else
            //        {
            //            worksheet.Cells[k * (rows + 4) + 1, 1].Value = ($"{p3} {pairs[2][k - 1].Item1}-{pairs[2][k - 1].Item2}");
            //            worksheet.Cells[k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
            //        }
            //        int row1 = k * (rows + 4) + 1;
            //        for (int j = 0; j < cols; j++)
            //        {
            //            if (j == 0)
            //            {
            //                worksheet.Cells[row1 + 1, j + 2].Value = ($"<{pairs[1][0].Item1}");
            //            }
            //            else if (j == cols - 1)
            //            {
            //                worksheet.Cells[row1 + 1, j + 2].Value = ($">{pairs[1][j - 2].Item2}");
            //            }
            //            else
            //            {
            //                worksheet.Cells[row1 + 1, j + 2].Value = ($"{pairs[1][j - 1].Item1}-{pairs[1][j - 1].Item2}");
            //            }
            //        }
            //        worksheet.Cells[row1 + 1, cols + 2].Value = "Sum";
            //        for (int i = 0; i < rows; i++)
            //        {
            //            double rowstotal = 0;

            //            if (i == 0)
            //            {
            //                worksheet.Cells[row1 + 2 + i, 1].Value = ($"<{pairs[0][0].Item1}");
            //            }
            //            else if (i == rows - 1)
            //            {
            //                worksheet.Cells[row1 + 2 + i, 1].Value = ($">{pairs[0][i - 2].Item2}");
            //            }
            //            else
            //            {
            //                worksheet.Cells[row1 + 2 + i, 1].Value = ($"{pairs[0][i - 1].Item1}-{pairs[0][i - 1].Item2}");
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
            //        worksheet.Cells[nextRows + k * (rows + 4) + 1, 1].Value = ($"{p3} <{pairs[2][0].Item1}");
            //        worksheet.Cells[nextRows + k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
            //    }
            //    else if (k == depth - 1)
            //    {
            //        worksheet.Cells[nextRows + k * (rows + 4) + 1, 1].Value = ($"{p3} <{pairs[2][0].Item1}");
            //        worksheet.Cells[nextRows + k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
            //    }
            //    else
            //    {
            //        worksheet.Cells[nextRows + k * (rows + 4) + 1, 1].Value = ($"{p3} {pairs[2][k - 1].Item1}-{pairs[2][k - 1].Item2}");
            //        worksheet.Cells[nextRows + k * (rows + 4) + 2, 1].Value = ($"{p1}/{p2}");
            //    }
            //    int row1 = k * (rows + 4) + 1;
            //    for (int j = 0; j < cols; j++)
            //    {
            //        if (j == 0)
            //        {
            //            worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($"<{pairs[1][0].Item1}");
            //        }
            //        else if (j == cols - 1)
            //        {
            //            worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($">{pairs[1][j - 2].Item2}");
            //        }
            //        else
            //        {
            //            worksheet.Cells[nextRows + row1 + 1, j + 2].Value = ($"{pairs[1][j - 1].Item1}-{pairs[1][j - 1].Item2}");
            //        }
            //    }
            //    worksheet.Cells[nextRows + row1 + 1, cols + 2].Value = "Sum";
            //    for (int i = 0; i < rows; i++)
            //    {
            //        double rowstotal = 0;

            //        if (i == 0)
            //        {
            //            worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($"<{pairs[0][0].Item1}");
            //        }
            //        else if (i == rows - 1)
            //        {
            //            worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($">{pairs[0][i - 2].Item2}");
            //        }
            //        else
            //        {
            //            worksheet.Cells[nextRows + row1 + 2 + i, 1].Value = ($"{pairs[0][i - 1].Item1}-{pairs[0][i - 1].Item2}");
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
            //        writer.WriteLine($"{p4} <{pairs[3][0].Item1},");
            //    }
            //    else if(t == time - 1)
            //    {
            //        writer.WriteLine($"{p4} >{pairs[3][t-2].Item2},");
            //    }
            //    else
            //    {
            //        writer.WriteLine($"{p4} {pairs[3][t-1].Item1}-{pairs[3][t-1].Item2},");
            //    }

            //    for (int k = 0; k < depth; k++)
            //    {
            //        if (k == 0)
            //        {
            //            writer.WriteLine($"{p3} <{pairs[2][0].Item1},");
            //        }
            //        else if (k == depth - 1)
            //        {
            //            writer.WriteLine($"{p3} >{pairs[2][k-2].Item2},");
            //        }
            //        else
            //        {
            //            writer.WriteLine($"{p3} {pairs[2][k-1].Item1}-{pairs[2][k-1].Item2},");
            //        }


            //        writer.Write(" ");
            //        for (int j = 0; j < cols; j++)
            //        {
            //            if (j == 0)
            //            {
            //                writer.Write($"{p1}/{p2},<{pairs[1][j].Item2},");
            //            }
            //            else if(j == cols - 1)
            //            {
            //                writer.Write($">{pairs[1][j-2].Item2},");
            //            }
            //            else
            //            {
            //                writer.Write($"{pairs[1][j-1].Item1}-{pairs[1][j-1].Item2},");
            //            }

            //        }
            //        writer.WriteLine();

            //        for (int i = 0; i < rows; i++)
            //        {
            //            if (i == 0)
            //            {
            //                writer.Write($"<{pairs[0][i].Item2},");
            //            }
            //            else if (i == rows-1)
            //            {
            //                writer.Write($">{pairs[0][i-2].Item2},");
            //            }
            //            else
            //            {
            //                writer.Write($"{pairs[0][i-1].Item1}-{pairs[0][i - 1].Item2},");
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

        List<List<(double, double)>> GeneratePermutations(List<(double, double)>[] pairs)
        {
            List<List<(double, double)>> result = new List<List<(double, double)>>();

            // 使用递归生成全排列
            GeneratePermutationsRecursive(pairs, 0, new List<(double, double)>(), result);

            return result;
        }

        void GeneratePermutationsRecursive(List<(double, double)>[] pairs, int index, List<(double, double)> current, List<List<(double, double)>> result)
        {
            if (index == pairs.Length)
            {
                result.Add(new List<(double, double)>(current));
                return;
            }

            foreach (var pair in pairs[index])
            {
                current.Add(pair);
                GeneratePermutationsRecursive(pairs, index + 1, current, result);
                current.RemoveAt(current.Count - 1);
            }
        }

        void SaveToCsv(List<List<(double, double)>> data, StreamWriter writer, int[] col)
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
                    line[col[0]] = row[0].Item1.ToString();
                    line[col[0]+1] = row[0].Item2.ToString();
                }
                if (row.Count > 1)
                {
                    line[col[1]] = row[1].Item1.ToString();
                    line[col[1]+1] = row[1].Item2.ToString();
                }
                if (row.Count > 2)
                {
                    line[col[2]] = row[2].Item1.ToString();
                    line[col[2]+1] = row[2].Item2.ToString();
                }
                if (row.Count > 3)
                {
                    line[col[3]] = row[3].Item1.ToString();
                    line[col[3]+1] = row[3].Item2.ToString();
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

            // 获取文件名
            string fileName = BinName.Text + ".csv";

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
                            throw new ArgumentException("Dimension must be between 1 and 4.");
                        }

                        double[] minValues = new double[dimension];
                        double[] stepSizes = new double[dimension];
                        int[] counts = new int[dimension];
                        int[] col = new int[dimension];

                        if (dimension >= 1)
                        {
                            if (dict.ContainsKey(this.para1.Text.ToUpper())) col[0] = dict[this.para1.Text.ToUpper()];
                            minValues[0] = double.Parse(this.para1min.Text);
                            stepSizes[0] = double.Parse(this.para1rta.Text);
                            counts[0] = int.Parse(this.para1num.Text);
                        }
                        if (dimension >= 2)
                        {
                            if (dict.ContainsKey(this.para2.Text.ToUpper())) col[1] = dict[this.para2.Text.ToUpper()];
                            minValues[1] = double.Parse(this.para2min.Text);
                            stepSizes[1] = double.Parse(this.para2rta.Text);
                            counts[1] = int.Parse(this.para2num.Text);
                        }
                        if (dimension >= 3)
                        {
                            if (dict.ContainsKey(this.para3.Text.ToUpper())) col[2] = dict[this.para3.Text.ToUpper()];
                            minValues[2] = double.Parse(this.para3min.Text);
                            stepSizes[2] = double.Parse(this.para3rta.Text);
                            counts[2] = int.Parse(this.para3num.Text);
                        }
                        if (dimension >= 4)
                        {
                            if (dict.ContainsKey(this.para4.Text.ToUpper())) col[3] = dict[this.para4.Text.ToUpper()];
                            minValues[3] = double.Parse(this.para4min.Text);
                            stepSizes[3] = double.Parse(this.para4rta.Text);
                            counts[3] = int.Parse(this.para4num.Text);
                        }

                        // 存储生成的pair数组
                        List<(double, double)>[] pairs = new List<(double, double)>[dimension];

                        // 生成pair数组
                        for (int i = 0; i < dimension; i++)
                        {
                            pairs[i] = new List<(double, double)>();
                            for (int j = 0; j < counts[i]; j++)
                            {
                                double first = minValues[i] + j * stepSizes[i];
                                double second = first + stepSizes[i];
                                pairs[i].Add((first, second));
                            }
                        }

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
