using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using static OfficeOpenXml.ExcelErrorValue;
using System.Threading.Tasks;
using System.Reflection;

namespace 产出分布计算
{
    public partial class Page4 : Page
    {

        private readonly object parameterlockObject = new object();
        private readonly Dictionary<string, int> dict2 = new Dictionary<string, int>
    {
        {"TEST", 0}, {"BIN", 1}, {"VF1", 2}, {"VF2", 3}, {"VF3", 4}, {"VF4", 5}, {"VF5", 6}, {"VF6", 7}, {"DVF", 8},
        {"VF", 9}, {"VFD", 10}, {"VZ1", 11}, {"VZ2", 12}, {"IR", 13}, {"LOP1", 14}, {"LOP2", 15}, {"LOP3", 16},
        {"WLP1", 17}, {"WLD1", 18}, {"WLC1", 19}, {"HW1", 20}, {"PURITY1", 21}, {"X1", 22}, {"Y1", 23}, {"Z1", 24},
        {"ST1", 25}, {"INT1", 26}, {"WLP2", 27}, {"WLD2", 28}, {"WLC2", 29}, {"HW2", 30}, {"PURITY2", 31}, {"DVF1", 32},
        {"DVF2", 33}, {"INT2", 34}, {"ST2", 35}, {"VF7", 36}, {"VF8", 37}, {"IR3", 38}, {"IR4", 39}, {"IR5", 40}, {"IR6", 41},
        {"VZ3", 42}, {"VZ4", 43}, {"VZ5", 44}, {"IF", 45}, {"IF1", 46}, {"IF2", 47}, {"ESD1", 48}, {"ESD2", 49}, {"IR1", 50},
        {"IR2", 51}, {"ESD1PASS", 52}, {"ESD2PASS", 53}, {"PosX", 54}, {"PosY", 55}
    };

        public Page4()
        {
            InitializeComponent();
            this.KeepAlive = true;
        }

        private async void ApplyChanges_Click(object sender, RoutedEventArgs e)
        {
            List<Task> tasks = new List<Task>();
            List<int> listIndex = new List<int>();
            int[] cols = new int[3];
            int[] cols2 = new int[3];

            if (p1.IsChecked == true)
            {
                listIndex.Add(0);

                if (!dict2.TryGetValue(TextBox1.Text.ToUpper(), out int columnNumber1))
                {
                    MessageBox.Show("Invalid column name for parameter 1.");
                }

                if (!dict2.TryGetValue(TextBox12.Text.ToUpper(), out int columnNumber12))
                {
                    MessageBox.Show("Invalid column name for parameter 1.");
                }
                cols[0] = columnNumber1;
                cols2[0] = columnNumber12;
            }

            if (p2.IsChecked == true)
            {
                listIndex.Add(1);
                if (!dict2.TryGetValue(TextBox2.Text.ToUpper(), out int columnNumber2))
                {
                    MessageBox.Show("Invalid column name for parameter 2.");
                }

                if (!dict2.TryGetValue(TextBox22.Text.ToUpper(), out int columnNumber22))
                {
                    MessageBox.Show("Invalid column name for parameter 2.");
                }

                cols[1] = columnNumber2;
                cols2[1] = columnNumber22;
            }

            if (p3.IsChecked == true)
            {
                listIndex.Add(2);
                if (!dict2.TryGetValue(TextBox3.Text.ToUpper(), out int columnNumber3))
                {
                    MessageBox.Show("Invalid column name for parameter 3.");
                }
                if (!dict2.TryGetValue(TextBox32.Text.ToUpper(), out int columnNumber32))
                {
                    MessageBox.Show("Invalid column name for parameter 3.");
                }

                cols[2] = columnNumber3;
                cols2[2] = columnNumber32;
            }

            var openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "CSV files (*.csv)|*.csv"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (var fileName in openFileDialog.FileNames)
                {
                    tasks.Add(Task.Run(() => ApplyChangesToFile(fileName, cols, cols2, listIndex)));
                }

                await Task.WhenAll(tasks);
                MessageBox.Show("Changes applied successfully.");
            }
        }

        private async Task ApplyChangesToFile(string fileName, int[] columnNumber, int[] columnNumber2, List<int> listIndex)
        {
            string tempFileName = Path.GetTempFileName();
            List<string> linesBuffer = new List<string>();
            int bufferSize = 500; // 每次批量写入 300 行
            int[] target = new int[3];
            int[] tempcol = new int[3];
            int[] tempcol2 = new int[3];

            target[0] = dict2["DVF"];
            target[1] = dict2["DVF1"];
            target[2] = dict2["DVF2"];

            using (var reader = new StreamReader(fileName))
            using (var writer = new StreamWriter(tempFileName, false, Encoding.UTF8))
            {
                string line;
                while ((line = await reader.ReadLineAsync()) != null)
                {
                    var lineValues = line.Split(',');

                    foreach (int index in listIndex)
                    {
                        if (lineValues.Length >= 56 && lineValues[0] == "TEST")
                        {
                            if (lineValues[1] == "BIN1" && lineValues[2] == "BIN2")
                            {
                                target[index]++;
                                tempcol[index] = columnNumber[index] + 1;
                                tempcol2[index] = columnNumber2[index] + 1;
                            }
                            else
                            {
                                tempcol[index] = columnNumber[index];
                                tempcol2[index] = columnNumber2[index];
                            }
                        }


                        string firstValue = lineValues[0];
                        if (Regex.IsMatch(firstValue, @"^\d+$") && lineValues.Length >= 56)
                        {

                            double Value1 = !string.IsNullOrEmpty(lineValues[tempcol[index]]) ? Convert.ToDouble(lineValues[tempcol[index]]) : 0;
                            double Value2 = !string.IsNullOrEmpty(lineValues[tempcol2[index]]) ? Convert.ToDouble(lineValues[tempcol2[index]]) : 0;
                            Value1 = Math.Round(Value1, 6);
                            Value2 = Math.Round(Value2, 6);
                            double newValue = Math.Round(Value1 - Value2, 6);

                            lineValues[target[index]] = newValue.ToString();
                        }
                    }

                    linesBuffer.Add(string.Join(",", lineValues));

                    if (linesBuffer.Count >= bufferSize)
                    {
                        await writer.WriteLineAsync(string.Join(Environment.NewLine, linesBuffer));
                        linesBuffer.Clear();
                    }
                }

                // 写入剩余的行
                if (linesBuffer.Count > 0)
                {
                    await writer.WriteLineAsync(string.Join(Environment.NewLine, linesBuffer));
                }
            }

            File.Delete(fileName);
            File.Move(tempFileName, fileName);

            await Dispatcher.InvokeAsync(() =>
            {
                fileListBox.Items.Add(fileName);
                fileListBox.ScrollIntoView(fileListBox.Items[fileListBox.Items.Count - 1]);
            });
        }

    }
}