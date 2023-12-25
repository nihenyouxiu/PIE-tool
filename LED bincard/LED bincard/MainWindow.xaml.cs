using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace WpfApp
{

    public partial class MainWindow : Window
    {
        int ParaIdx1 = 0;
        int ParaIdx2 = 0;
        int ParaIdx3 = 0;
        int ParaIdx4 = 0;

        double WLPMinFloat = 0;
        double WLPMaxFloat = 0;
        double WLPStepFloat= 0;
        double LOPMinFloat = 0;
        double LOPMaxFloat = 0;
        double LOPStepFloat= 0;
        double VF1MinFloat = 0;
        double VF1MaxFloat = 0;
        double VF1StepFloat= 0;
        double VF2MinFloat = 0;
        double VF2MaxFloat = 0;
        double VF2StepFloat= 0;

        int isWLPSubmit = 0;
        int isLOPSubmit = 0;
        int isVF1Submit = 0;
        int isVF2Submit = 0;

        double totalChipNum = 0;
        double BinCardChipNum = 0;
        double NgChipNum = 0;

        public MainWindow()
        {
            InitializeComponent();
        }
        private int CompareDouble(double WLP1Value, double VF1Value, double LOP1Value, double VF2Value)
        {
            if (isWLPSubmit == 1 && (WLP1Value < WLPMinFloat || WLP1Value > WLPMaxFloat))
            {
                return 1;
            } 
            if(isLOPSubmit == 1 && (LOP1Value < LOPMinFloat || LOP1Value > LOPMaxFloat))
            {
                return 2;
            }
            
            if (isVF1Submit == 1 && (VF1Value < VF1MinFloat || VF1Value > VF1MaxFloat))
            {
                return 3;
            }
            if(isVF2Submit == 1 && (VF2Value < VF2MinFloat || VF2Value > VF2MaxFloat))
            {
                return 4;
            }

            return 0;
        }

        string directoryPath = null;
        int WLPArrayCnt = 1;
        int LOPArrayCnt = 1;
        int VF1ArrayCnt = 1;
        int VF2ArrayCnt = 1;
        private void OnSelectCsvClick(object sender, RoutedEventArgs e)
        {
            // 使用OpenFileDialog选择多个CSV文件
            string lastCellValue = null;
            
            if (isWLPSubmit == 1)
            {
                ParaIdx1 = (int)OnIndexCheckClick(Para1Index.Text);
                WLPMinFloat = OnCheckClick(WLPMin.Text);
                WLPMaxFloat = OnCheckClick(WLPMax.Text);
                WLPStepFloat = OnCheckClick(WLPStep.Text);
                WLPArrayCnt = (int)Math.Floor((WLPMaxFloat - WLPMinFloat) / WLPStepFloat);
            }

            if (isLOPSubmit == 1)
            {
                ParaIdx2 = (int)OnIndexCheckClick(Para2Index.Text);
                LOPMinFloat = OnCheckClick(LopMin.Text);
                LOPMaxFloat = OnCheckClick(LopMax.Text);
                LOPStepFloat = OnCheckClick(LopStep.Text);
                LOPArrayCnt = (int)Math.Floor((LOPMaxFloat - LOPMinFloat) / LOPStepFloat);
            }

            if (isVF1Submit == 1)
            {
                ParaIdx3 = (int)OnIndexCheckClick(Para3Index.Text);
                VF1MinFloat = OnCheckClick(VF1Min.Text);
                VF1MaxFloat = OnCheckClick(VF1Max.Text);
                VF1StepFloat = OnCheckClick(VF1Step.Text); 
                VF1ArrayCnt = (int)Math.Floor((VF1MaxFloat - VF1MinFloat) / VF1StepFloat);
            }

            if (isVF2Submit == 1)
            {
                ParaIdx4 = (int)OnIndexCheckClick(Para4Index.Text);
                VF2MinFloat = OnCheckClick(VF2Min.Text);
                VF2MaxFloat = OnCheckClick(VF2Max.Text);
                VF2StepFloat = OnCheckClick(VF2Step.Text);
                VF2ArrayCnt = (int)Math.Floor((VF2MaxFloat - VF2MinFloat) / VF2StepFloat);
            }

            int[,,,] fourDimensionalArray = new int[WLPArrayCnt, LOPArrayCnt, VF1ArrayCnt, VF2ArrayCnt]; // 四维数组
            
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {

                // 读取并显示每个CSV文件的内容
                foreach (string csvFilePath in openFileDialog.FileNames)
                {
 
                    // 构造输出文件路径，在输入文件路径下添加 "out.csv" 后缀
                    string outputFileName = Path.GetFileNameWithoutExtension(csvFilePath) + "_out.csv";
                    string outputFile = Path.Combine(Path.GetDirectoryName(csvFilePath), outputFileName);
                    directoryPath = Path.GetDirectoryName(csvFilePath);
                    List<string[]> csvData = ReadCsvFile(csvFilePath, csvFilePath, lastCellValue, outputFile, fourDimensionalArray);

                }
            }
            // 构造新文件的完整路径
            string newFileName = "result_out.csv";
            string newFilePath = Path.Combine(directoryPath, newFileName);

            using (StreamWriter writer = new StreamWriter(newFilePath,false, Encoding.UTF8))
            {

                int index = 0;
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "CSV 文件 (*.csv)|*.csv";

                writer.WriteLine("Index,Para1,Para2,Para3,Para4,Count,Percentage"); // 这里假设保存的是从1开始的整数序列，可以根据实际需求修改


                for (int WLPArrayIdx = 0; WLPArrayIdx < WLPArrayCnt; WLPArrayIdx++)
                {
                    for (int LOPArrayIdx = 0; LOPArrayIdx < LOPArrayCnt; LOPArrayIdx++)
                    {
                        for (int VF1ArrayIdx = 0; VF1ArrayIdx < VF1ArrayCnt; VF1ArrayIdx++)
                        {
                            for (int VF2ArrayIdx = 0; VF2ArrayIdx < VF2ArrayCnt; VF2ArrayIdx++)
                            {
                                index++;
                                double tmp = fourDimensionalArray[WLPArrayIdx, LOPArrayIdx, VF1ArrayIdx, VF2ArrayIdx];
                                string lineValue = String.Concat(index.ToString(), ",",
                                    (WLPMinFloat + (double)WLPArrayIdx * WLPStepFloat).ToString(), "~", Math.Min(WLPMaxFloat, (WLPMinFloat + (double)(WLPArrayIdx + 1) * WLPStepFloat)).ToString(), ",",
                                    (LOPMinFloat + (double)LOPArrayIdx * LOPStepFloat).ToString(), "~", Math.Min(LOPMaxFloat, (LOPMinFloat + (double)(LOPArrayIdx + 1) * LOPStepFloat)).ToString(), ",",
                                    (VF1MinFloat + (double)VF1ArrayIdx * VF1StepFloat).ToString(), "~", Math.Min(VF1MaxFloat, (VF1MinFloat + (double)(VF1ArrayIdx + 1) * VF1StepFloat)).ToString(), ",",
                                    (VF2MinFloat + (double)VF2ArrayIdx * VF2StepFloat).ToString(), "~", Math.Min(VF2MaxFloat, (VF2MinFloat + (double)(VF2ArrayIdx + 1) * VF2StepFloat)).ToString(), ",",
                                    tmp, ",", (tmp / (double)totalChipNum)*100,"%");
                                writer.WriteLine(lineValue);
                            }
                        }
                    }
                }
                string newlineValue = String.Concat(",,,,,", totalChipNum);
                writer.WriteLine(newlineValue);
            }


            MessageBox.Show("CSV 文件读取并保存成功！");
        }

        private double OnCheckClick(string userInput)
        {

            // 判断文本框是否为空
            if (string.IsNullOrWhiteSpace(userInput))
            {
                MessageBox.Show("请输入整数");
                return -1;
            }

            // 尝试将输入的字符串转换为浮点型
            if (double.TryParse(userInput, out double floatValue))
            {
               return floatValue;
            }
            else
            {
                return -1;
            }
        }

        private double OnIndexCheckClick(string userInput)
        {

            // 判断文本框是否为空
            if (string.IsNullOrWhiteSpace(userInput))
            {
                MessageBox.Show("请输入浮点数");
                return -1;
            }

            // 尝试将输入的字符串转换为浮点型
            if (int.TryParse(userInput, out int intValue))
            {
                return intValue;
            }
            else
            {
                return -1;
            }
        }

        private void OnSubmitClick(object sender, RoutedEventArgs e)
        {
            // 获取选中的选项
            List<string> selectedOptions = new List<string>();


            if (checkBoxWLP.IsChecked == true)
            {
                selectedOptions.Add("Para1");
                isWLPSubmit = 1;
            }

            if (checkBoxLOP.IsChecked == true)
            {
                selectedOptions.Add("Para2");
                isLOPSubmit = 1;
            }

            if (checkBoxVF1.IsChecked == true)
            {
                selectedOptions.Add("Para3");
                isVF1Submit = 1;
            }

            if (checkBoxVF2.IsChecked == true)
            {
                selectedOptions.Add("Para4");
                isVF2Submit = 1;
            }

            // 处理选中的选项，可以根据实际需求进行处理
            MessageBox.Show("选中的选项：" + string.Join(", ", selectedOptions));
        }

        int cnt = 0;

        private void SaveArrayAsCsv(StreamWriter writer,string filePath,double VF1Value, double VF2Value, double LOP1Value, double WLP1Value)
        {
            writer.WriteLine(VF1Value.ToString(),",", VF2Value.ToString(), ",", LOP1Value.ToString(), ",", WLP1Value.ToString()); // 这里假设保存的是从1开始的整数序列，可以根据实际需求修改
        }

        private List<string[]> ReadCsvFile(string filename, string filePath, string lastCellValue, string outputFile, int[,,,] fourDimensionalArray)             
        {
            List<string[]> csvData = new List<string[]>();
            double tmp = 0;
            try
            {
                // 读取CSV文件的从第19行开始的内容
                List<string> lines = new List<string>();
                string line;
                using (StreamReader reader = new StreamReader(filePath, Encoding.UTF8))
                using (StreamWriter writer = new StreamWriter(outputFile,false, Encoding.UTF8))
                {

                    string outputText = null;
                    SaveFileDialog saveFileDialog = new SaveFileDialog();
                    saveFileDialog.Filter = "CSV 文件 (*.csv)|*.csv";
                    for (int i = 0; i < 18; i++)
                    {
                        // 前18行;
                        line = reader.ReadLine();

                        if (line.StartsWith("model", StringComparison.OrdinalIgnoreCase))
                        {
                            // 输出匹配行的两个逗号后的内容
                            string[] values = line.Split(',');
                            if (values.Length >= 3)
                            {
                                outputText = $"{values[2].Substring(0, Math.Min(values[2].Length, 8))}";
                                string lineValue = String.Concat( filename);
                                writer.WriteLine(lineValue);
                                lineValue = String.Concat("Model,", outputText);
                                writer.WriteLine(lineValue);
                                writer.WriteLine("Index,Para1Index,Para2Index,Para3Index,Para4Index,Para1Value,Para2Value,Para3Value,Para4Value,NgIndex,Result"); // 这里假设保存的是从1开始的整数序列，可以根据实际需求修改
                            }
                        }
                    }
                    NgChipNum = 0;
                    BinCardChipNum = 0;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (double.TryParse(line.Split(',')[ParaIdx3], out double VF1Value) && double.TryParse(line.Split(',')[ParaIdx4], out double VF2Value) &&
                            double.TryParse(line.Split(',')[ParaIdx2], out double LOP1Value) && double.TryParse(line.Split(',')[ParaIdx1], out double WLP1Value))
                        {
                            int WLPIndex = 0;
                            int LOPIndex = 0;
                            int VF1Index = 0;
                            int VF2Index = 0;

                            int tmep = CompareDouble(WLP1Value, VF1Value, LOP1Value, VF2Value);
                            if(tmep == 0)
                                {
                        
                                    if (isWLPSubmit == 1)
                                    {
                                        WLPIndex = (int)Math.Floor((WLP1Value - WLPMinFloat) / WLPStepFloat);
                                        if(WLP1Value == WLPMaxFloat)
                                        {
                                            WLPIndex = WLPArrayCnt - 1;
                                        }
                                    }
                        
                                    if (isLOPSubmit == 1)
                                    {
                        
                                        LOPIndex = (int)Math.Floor((LOP1Value - LOPMinFloat) / LOPStepFloat);
                                        if (LOP1Value == LOPMaxFloat)
                                        {
                                            LOPIndex = LOPArrayCnt - 1;
                                        }
                                    }
                        
                                    if (isVF1Submit == 1)
                                    {
                        
                                        VF1Index = (int)Math.Floor((VF1Value - VF1MinFloat) / VF1StepFloat);
                                        if (VF1Value == VF1MaxFloat)
                                        {
                                            VF1Index = VF1ArrayCnt - 1;
                                        }
                                    }
                        
                                    if (isVF2Submit == 1)
                                    {
                                        VF2Index = (int)Math.Floor((VF2Value - VF2MinFloat) / VF2StepFloat);
                                        if (VF2Value == VF2MaxFloat)
                                        {
                                            VF2Index = VF2ArrayCnt - 1;
                                        }
                                    }
                            
                        
                                    string lineValue = String.Concat(line.Split(',')[0], ",", WLPIndex, ",", LOPIndex, ",", VF1Index, ",", VF2Index, ",");
                                    //lineValue = String.Concat(lineValue, ",", WLPMaxFloat, ",", WLPMinFloat, ",");
                                    lineValue = String.Concat(lineValue, WLP1Value, ",", LOP1Value, ",", VF1Value, ",", VF2Value, ",", tmep, ",", "OK");
                                    BinCardChipNum++;
                                    writer.WriteLine(lineValue);
                                    fourDimensionalArray[WLPIndex, LOPIndex, VF1Index, VF2Index]++;
                                }
                            else
                          {
                              // 转换成功，doubleValue 包含转换后的 double 值
                              NgChipNum++;
                              string lineValue = String.Concat(line.Split(',')[0], ",", WLPIndex, ",", LOPIndex, ",", VF1Index, ",", VF2Index, ",");
                              //lineValue = String.Concat(lineValue, ",", WLPMaxFloat, ",", WLPMinFloat, ",");
                              lineValue = String.Concat(lineValue, WLP1Value, ",", LOP1Value, ",", VF1Value, ",", VF2Value, ",", tmep, ",", "Ng");
                              writer.WriteLine(lineValue);
                          }

                        }
                        else
                        {
                            // 转换失败，stringValue 不是有效的 double 表示 
                            MessageBox.Show($"Error convert to double type '{filePath}'");

                        }

                        lastCellValue = line.Split(',')[0];
                        if (double.TryParse(lastCellValue, out double doubleValue))
                        {
                            // doubleValue，intValue 包含转换后的整数值
                            tmp = doubleValue;
                        }
                        else
                        {
                            // 转换失败，stringValue 不是有效的整数表示
                        }
                    }
                    cnt++;
                    totalChipNum = totalChipNum + tmp;
                    outputText = String.Concat(cnt, " ", filename, "  ", "Model: ", outputText, " Num:", tmp.ToString(),",Ng Chip Num", NgChipNum.ToString()
                        , ",BinCard Chip Num", BinCardChipNum.ToString(), ",Total Chip Num:", totalChipNum.ToString());
                    AddResultToScrollViewer(outputText);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading CSV file '{filePath}': {ex.Message}");
            }

            return csvData;
        }

        private void AddResultToScrollViewer(string result)
        {
            // 创建一个文本块并添加到滚动控件中
            TextBlock textBlock = new TextBlock
            {
                Text = result,
                Margin = new Thickness(0, 5, 0, 5)
            };

            resultStackPanel.Children.Add(textBlock);
        }

        private void OnClearButtonClick(object sender, RoutedEventArgs e)
        {
            // 当按钮被点击时，清空所有文本框的内容
            Para1Index.Text = string.Empty;
            Para2Index.Text = string.Empty;
            Para3Index.Text = string.Empty;
            Para4Index.Text = string.Empty;

            WLPMin.Text = string.Empty;
            WLPMax.Text = string.Empty;
            WLPStep.Text = string.Empty;
            LopMin.Text = string.Empty;
            LopMax.Text = string.Empty;
            LopStep.Text = string.Empty;
            VF1Min.Text = string.Empty;
            VF1Max.Text = string.Empty;
            VF1Step.Text = string.Empty;
            VF2Min.Text = string.Empty;
            VF2Max.Text = string.Empty;
            VF2Step.Text = string.Empty;
        }

    }
}
