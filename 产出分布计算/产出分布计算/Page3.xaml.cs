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
    public partial class Page3 : Page
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

        public Page3()
        {
            InitializeComponent();
            this.KeepAlive = true;
        }

        private async void ApplyChanges_Click(object sender, RoutedEventArgs e)
        {
            List<Task> tasks = new List<Task>();
            List<int> listIndex= new List<int>();
            string[] operation = new string[3]; 
            int[] cols = new int[3];
            double[] values = new double[3];

            if (p1.IsChecked == true)
            {
                listIndex.Add(0);

                if (!dict2.TryGetValue(TextBox1.Text.ToUpper(), out int columnNumber1))
                {
                    MessageBox.Show("Invalid column name for parameter 1.");
                }

                string operation1 = (operationComboBox.SelectedItem as ComboBoxItem)?.Content as string;
                if (string.IsNullOrEmpty(operation1))
                {
                    MessageBox.Show("Please select an operation for parameter 1.");
                }

                if (!double.TryParse(valueTextBox1.Text, out double value1))
                {
                    MessageBox.Show("Please enter a valid value for parameter 1.");
                }
                cols[0] = columnNumber1;
                operation[0] = operation1;
                values[0] = value1;
            }

            if (p2.IsChecked == true)
            {
                listIndex.Add(1);
                if (!dict2.TryGetValue(TextBox2.Text.ToUpper(), out int columnNumber2))
                {
                    MessageBox.Show("Invalid column name for parameter 2.");
                }

                string operation2 = (operationComboBox2.SelectedItem as ComboBoxItem)?.Content as string;
                if (string.IsNullOrEmpty(operation2))
                {
                    MessageBox.Show("Please select an operation for parameter 2.");
                }

                if (!double.TryParse(valueTextBox2.Text, out double value2))
                {
                    MessageBox.Show("Please enter a valid value for parameter 2.");
                }
                cols[1] = columnNumber2;
                operation[1] = operation2;
                values[1] = value2;
            }

            if (p3.IsChecked == true)
            {
                listIndex.Add(2);
                if (!dict2.TryGetValue(TextBox3.Text.ToUpper(), out int columnNumber3))
                {
                    MessageBox.Show("Invalid column name for parameter 3.");
                }

                string operation3 = (operationComboBox3.SelectedItem as ComboBoxItem)?.Content as string;
                if (string.IsNullOrEmpty(operation3))
                {
                    MessageBox.Show("Please select an operation for parameter 3.");
                }

                if (!double.TryParse(valueTextBox3.Text, out double value3))
                {
                    MessageBox.Show("Please enter a valid value for parameter 3.");
                }
                cols[2] = columnNumber3;
                operation[2] = operation3;
                values[2] = value3;
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
                    tasks.Add(Task.Run(() => ApplyChangesToFile(fileName, cols, operation, values, listIndex)));
                }
                
                await Task.WhenAll(tasks);
                MessageBox.Show("Changes applied successfully.");
            }
        }

        private async Task ApplyChangesToFile(string fileName, int[] columnNumber, string[] operation, double[] value, List<int> listIndex)
        {
            string tempFileName = Path.GetTempFileName();
            List<string> linesBuffer = new List<string>();
            int bufferSize = 500; // 每次批量写入 300 行
            int[] tempcol = new int[3];

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
                                tempcol[index] = columnNumber[index] + 1;
                            else
                                tempcol[index] = columnNumber[index];
                        }

                        string firstValue = lineValues[0];
                        if (Regex.IsMatch(firstValue, @"^\d+$") && lineValues.Length >= 56)
                        {
                            if (operation[index] == "设定值")
                            {
                                lineValues[tempcol[index]] = value.ToString();
                            }
                            else if (double.TryParse(lineValues[tempcol[index]], out double currentValue))
                            {
                                double newValue = CalculateNewValue(currentValue, operation[index], value[index]);
                                lineValues[tempcol[index]] = newValue.ToString();
                            }
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
    }
}