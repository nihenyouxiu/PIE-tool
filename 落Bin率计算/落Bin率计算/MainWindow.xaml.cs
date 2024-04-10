using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using CsvHelper;
using CsvHelper.Configuration;
using System.Windows.Controls;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;
using System.Windows;

namespace 落Bin率计算
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    public class BinData
    {
        public double binIdx { get; set; }
        public double VF1Min { get; set; }
        public double VF1Max { get; set; }
        public double VF2Min { get; set; }
        public double VF2Max { get; set; }
        public double VF3Min { get; set; }
        public double VF3Max { get; set; }
        public double VF4Min { get; set; }
        public double VF4Max { get; set; }
        public double VZ1Min { get; set; }
        public double VZ1Max { get; set; }
        public double IRMin { get; set; }
        public double IRMax { get; set; }
        public double HW1Min { get; set; }
        public double HW1Max { get; set; }
        public double LOP1Min { get; set; }
        public double LOP1Max { get; set; }
        public double WLP1Min { get; set; }
        public double WLP1Max { get; set; }
        public double WLD1Min { get; set; }
        public double WLD1Max { get; set; }
        public double IR1Min { get; set; }
        public double IR1Max { get; set; }
        public double VFDMin { get; set; }
        public double VFDMax { get; set; }
        public double DVFMin { get; set; }
        public double DVFMax { get; set; }
        public double IR2Min { get; set; }
        public double IR2Max { get; set; }
        public double WLC1Min { get; set; }
        public double WLC1Max { get; set; }
        public double VF5Min { get; set; }
        public double VF5Max { get; set; }
        public double VF6Min { get; set; }
        public double VF6Max { get; set; }
        public double VF7Min { get; set; }
        public double VF7Max { get; set; }
        public double VF8Min { get; set; }
        public double VF8Max { get; set; }
        public double DVF1Min { get; set; }
        public double DVF1Max { get; set; }
        public double DVF2Min { get; set; }
        public double DVF2Max { get; set; }
        public double VZ2Min { get; set; }
        public double VZ2Max { get; set; }
        public double VZ3Min { get; set; }
        public double VZ3Max { get; set; }
        public double VZ4Min { get; set; }
        public double VZ4Max { get; set; }
        public double VZ5Min { get; set; }
        public double VZ5Max { get; set; }
        public double IR3Min { get; set; }
        public double IR3Max { get; set; }
        public double IR4Min { get; set; }
        public double IR4Max { get; set; }
        public double IR5Min { get; set; }
        public double IR5Max { get; set; }
        public double IR6Min { get; set; }
        public double IR6Max { get; set; }
        public double IFMin { get; set; }
        public double IFMax { get; set; }
        public double IF1Min { get; set; }
        public double IF1Max { get; set; }
        public double IF2Min { get; set; }
        public double IF2Max { get; set; }
        public double LOP2Min { get; set; }
        public double LOP2Max { get; set; }
        public double WLP2Min { get; set; }
        public double WLP2Max { get; set; }
        public double WLD2Min { get; set; }
        public double WLD2Max { get; set; }
        public double HW2Min { get; set; }
        public double HW2Max { get; set; }
        public double WLC2Min { get; set; }
        public double WLC2Max { get; set; }
    }

    public class Chip
    {
        public double TEST { get; set; }
        public double BIN { get; set; }
        public double VF1 { get; set; }
        public double VF2 { get; set; }
        public double VF3 { get; set; }
        public double VF4 { get; set; }
        public double VF5 { get; set; }
        public double VF6 { get; set; }
        public double DVF { get; set; }
        public double VF { get; set; }
        public double VFD { get; set; }
        public double VZ1 { get; set; }
        public double VZ2 { get; set; }
        public double IR { get; set; }
        public double LOP1 { get; set; }
        public double LOP2 { get; set; }
        public double LOP3 { get; set; }
        public double WLP1 { get; set; }
        public double WLD1 { get; set; }
        public double WLC1 { get; set; }
        public double HW1 { get; set; }
        public double WLP2 { get; set; }
        public double WLD2 { get; set; }
        public double WLC2 { get; set; }
        public double HW2 { get; set; }
        public double DVF1 { get; set; }
        public double DVF2 { get; set; }
        public double VF7 { get; set; }
        public double VF8 { get; set; }
        public double IR3 { get; set; }
        public double IR4 { get; set; }
        public double IR5 { get; set; }
        public double IR6 { get; set; }
        public double VZ3 { get; set; }
        public double VZ4 { get; set; }
        public double VZ5 { get; set; }
        public double IF { get; set; }
        public double IF1 { get; set; }
        public double IF2 { get; set; }
        public double IR1 { get; set; }
        public double IR2 { get; set; }

        public override string ToString()
        {
            return $"TEST: {TEST}, Bin: {BIN}, VF1: {VF1}, VF2: {VF2}, VF3: {VF3}, VF4: {VF4}, VF5: {VF5}, VF6: {VF6}, " +
                   $"DVF: {DVF}, VF: {VF}, VFD: {VFD}, VZ1: {VZ1}, VZ2: {VZ2}, IR: {IR}, " +
                   $"LOP1: {LOP1}, LOP2: {LOP2}, LOP3: {LOP3}, WLP1: {WLP1}, WLD1: {WLD1}, " +
                   $"WLC1: {WLC1}, HW1: {HW1}, WLP2: {WLP2}, WLD2: {WLD2}, WLC2: {WLC2}, " +
                   $"HW2: {HW2}, DVF1: {DVF1}, DVF2: {DVF2}, VF7: {VF7}, VF8: {VF8}, IR3: {IR3}, " +
                   $"IR4: {IR4}, IR5: {IR5}, IR6: {IR6}, VZ3: {VZ3}, VZ4: {VZ4}, VZ5: {VZ5}, " +
                   $"IF: {IF}, IF1: {IF1}, IF2: {IF2}, IR1: {IR1}, IR2: {IR2.ToString()}"; // 将 IR2 转换为字符串
        }
    }

    public partial class MainWindow : Window
    {
        List<BinData> binDataList;


        public MainWindow()
        {
            InitializeComponent();
        }

        double vf1Min = -1000000;
        double vf1Max = -1000000;
        double vf2Min = -1000000;
        double vf2Max = -1000000;
        double vf3Min = -1000000;
        double vf3Max = -1000000;
        double vf4Min = -1000000;
        double vf4Max = -1000000;
        double vz1Min = -1000000;
        double vz1Max = -1000000;
        double irMin = -1000000;
        double irMax = -1000000;
        double hw1Min = -1000000;
        double hw1Max = -1000000;
        double lop1Min = -1000000;
        double lop1Max = -1000000;
        double wlp1Min = -1000000;
        double wlp1Max = -1000000;
        double wld1Min = -1000000;
        double wld1Max = -1000000;
        double ir1Min = -1000000;
        double ir1Max = -1000000;
        double vfdMin = -1000000;
        double vfdMax = -1000000;
        double dvfMin = -1000000;
        double dvfMax = -1000000;
        double ir2Min = -1000000;
        double ir2Max = -1000000;
        double wlc1Min = -1000000;
        double wlc1Max = -1000000;
        double vf5Min = -1000000;
        double vf5Max = -1000000;
        double vf6Min = -1000000;
        double vf6Max = -1000000;
        double vf7Min = -1000000;
        double vf7Max = -1000000;
        double vf8Min = -1000000;
        double vf8Max = -1000000;
        double dvf1Min = -1000000;
        double dvf1Max = -1000000;
        double dvf2Min = -1000000;
        double dvf2Max = -1000000;
        double vz2Min = -1000000;
        double vz2Max = -1000000;
        double vz3Min = -1000000;
        double vz3Max = -1000000;
        double vz4Min = -1000000;
        double vz4Max = -1000000;
        double vz5Min = -1000000;
        double vz5Max = -1000000;
        double ir3Min = -1000000;
        double ir3Max = -1000000;
        double ir4Min = -1000000;
        double ir4Max = -1000000;
        double ir5Min = -1000000;
        double ir5Max = -1000000;
        double ir6Min = -1000000;
        double ir6Max = -1000000;
        double ifMin = -1000000;
        double ifMax = -1000000;
        double if1Min = -1000000;
        double if1Max = -1000000;
        double if2Min = -1000000;
        double if2Max = -1000000;
        double lop2Min = -1000000;
        double lop2Max = -1000000;
        double wlp2Min = -1000000;
        double wlp2Max = -1000000;
        double wld2Min = -1000000;
        double wld2Max = -1000000;
        double hw2Min = -1000000;
        double hw2Max = -1000000;
        double wlc2Min = -1000000;
        double wlc2Max = -1000000;

        void getMaxMin()
        {
            // 获取所有属性的最小值和最大值
            vf1Min = binDataList.Min(data => data.VF1Min);
            vf1Max = binDataList.Max(data => data.VF1Max);

            vf2Min = binDataList.Min(data => data.VF2Min);
            vf2Max = binDataList.Max(data => data.VF2Max);

            vf3Min = binDataList.Min(data => data.VF3Min);
            vf3Max = binDataList.Max(data => data.VF3Max);

            vf4Min = binDataList.Min(data => data.VF4Min);
            vf4Max = binDataList.Max(data => data.VF4Max);

            vz1Min = binDataList.Min(data => data.VZ1Min);
            vz1Max = binDataList.Max(data => data.VZ1Max);

            irMin = binDataList.Min(data => data.IRMin);
            irMax = binDataList.Max(data => data.IRMax);

            hw1Min = binDataList.Min(data => data.HW1Min);
            hw1Max = binDataList.Max(data => data.HW1Max);

            lop1Min = binDataList.Min(data => data.LOP1Min);
            lop1Max = binDataList.Max(data => data.LOP1Max);

            wlp1Min = binDataList.Min(data => data.WLP1Min);
            wlp1Max = binDataList.Max(data => data.WLP1Max);

            wld1Min = binDataList.Min(data => data.WLD1Min);
            wld1Max = binDataList.Max(data => data.WLD1Max);

            ir1Min = binDataList.Min(data => data.IR1Min);
            ir1Max = binDataList.Max(data => data.IR1Max);

            vfdMin = binDataList.Min(data => data.VFDMin);
            vfdMax = binDataList.Max(data => data.VFDMax);

            dvfMin = binDataList.Min(data => data.DVFMin);
            dvfMax = binDataList.Max(data => data.DVFMax);

            ir2Min = binDataList.Min(data => data.IR2Min);
            ir2Max = binDataList.Max(data => data.IR2Max);

            wlc1Min = binDataList.Min(data => data.WLC1Min);
            wlc1Max = binDataList.Max(data => data.WLC1Max);

            vf5Min = binDataList.Min(data => data.VF5Min);
            vf5Max = binDataList.Max(data => data.VF5Max);

            vf6Min = binDataList.Min(data => data.VF6Min);
            vf6Max = binDataList.Max(data => data.VF6Max);

            vf7Min = binDataList.Min(data => data.VF7Min);
            vf7Max = binDataList.Max(data => data.VF7Max);

            vf8Min = binDataList.Min(data => data.VF8Min);
            vf8Max = binDataList.Max(data => data.VF8Max);

            dvf1Min = binDataList.Min(data => data.DVF1Min);
            dvf1Max = binDataList.Max(data => data.DVF1Max);

            dvf2Min = binDataList.Min(data => data.DVF2Min);
            dvf2Max = binDataList.Max(data => data.DVF2Max);

            vz2Min = binDataList.Min(data => data.VZ2Min);
            vz2Max = binDataList.Max(data => data.VZ2Max);

            vz3Min = binDataList.Min(data => data.VZ3Min);
            vz3Max = binDataList.Max(data => data.VZ3Max);

            vz4Min = binDataList.Min(data => data.VZ4Min);
            vz4Max = binDataList.Max(data => data.VZ4Max);

            vz5Min = binDataList.Min(data => data.VZ5Min);
            vz5Max = binDataList.Max(data => data.VZ5Max);

            ir3Min = binDataList.Min(data => data.IR3Min);
            ir3Max = binDataList.Max(data => data.IR3Max);

            ir4Min = binDataList.Min(data => data.IR4Min);
            ir4Max = binDataList.Max(data => data.IR4Max);

            ir5Min = binDataList.Min(data => data.IR5Min);
            ir5Max = binDataList.Max(data => data.IR5Max);

            ir6Min = binDataList.Min(data => data.IR6Min);
            ir6Max = binDataList.Max(data => data.IR6Max);

            ifMin = binDataList.Min(data => data.IFMin);
            ifMax = binDataList.Max(data => data.IFMax);

            if1Min = binDataList.Min(data => data.IF1Min);
            if1Max = binDataList.Max(data => data.IF1Max);

            if2Min = binDataList.Min(data => data.IF2Min);
            if2Max = binDataList.Max(data => data.IF2Max);

            lop2Min = binDataList.Min(data => data.LOP2Min);
            lop2Max = binDataList.Max(data => data.LOP2Max);

            wlp2Min = binDataList.Min(data => data.WLP2Min);
            wlp2Max = binDataList.Max(data => data.WLP2Max);

            wld2Min = binDataList.Min(data => data.WLD2Min);
            wld2Max = binDataList.Max(data => data.WLD2Max);

            hw2Min = binDataList.Min(data => data.HW2Min);
            hw2Max = binDataList.Max(data => data.HW2Max);

            wlc2Min = binDataList.Min(data => data.WLC2Min);
            wlc2Max = binDataList.Max(data => data.WLC2Max);
        }
        public class DataItem
        {
            public string Min { get; set; }
            public string Max { get; set; }
        }

        private void BinImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV files (*.csv)|*.csv";
            if (openFileDialog.ShowDialog() == true)
            {
                binDataList = new List<BinData>();

                using (var reader = new StreamReader(openFileDialog.FileName, Encoding.UTF8))
                {
                    // 跳过前7行
                    for (int i = 0; i < 7; i++)
                    {
                        reader.ReadLine();
                    }

                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        ListBoxItem item = new ListBoxItem();
                        string[] values = line.Split(',');

                        if (line.StartsWith("1"))
                        {
                            BinData binData = new BinData();
                            //将每个字段的值分配给相应的属性
                            binData.binIdx = !string.IsNullOrEmpty(values[2]) ? Convert.ToDouble(values[2]) : 0.0;
                            binData.VF1Min = !string.IsNullOrEmpty(values[4]) ? Convert.ToDouble(values[4]) : 0.0;
                            binData.VF1Max = !string.IsNullOrEmpty(values[5]) ? Convert.ToDouble(values[5]) : 0.0;
                            binData.VF2Min = !string.IsNullOrEmpty(values[6]) ? Convert.ToDouble(values[6]) : 0.0;
                            binData.VF2Max = !string.IsNullOrEmpty(values[7]) ? Convert.ToDouble(values[7]) : 0.0;
                            binData.VF3Min = !string.IsNullOrEmpty(values[8]) ? Convert.ToDouble(values[8]) : 0.0;
                            binData.VF3Max = !string.IsNullOrEmpty(values[9]) ? Convert.ToDouble(values[9]) : 0.0;
                            binData.VF4Min = !string.IsNullOrEmpty(values[10]) ? Convert.ToDouble(values[10]) : 0.0;
                            binData.VF4Max = !string.IsNullOrEmpty(values[11]) ? Convert.ToDouble(values[11]) : 0.0;
                            binData.VZ1Min = !string.IsNullOrEmpty(values[12]) ? Convert.ToDouble(values[12]) : 0.0;
                            binData.VZ1Max = !string.IsNullOrEmpty(values[13]) ? Convert.ToDouble(values[13]) : 0.0;
                            binData.IRMin = !string.IsNullOrEmpty(values[14]) ? Convert.ToDouble(values[14]) : 0.0;
                            binData.IRMax = !string.IsNullOrEmpty(values[15]) ? Convert.ToDouble(values[15]) : 0.0;
                            binData.HW1Min = !string.IsNullOrEmpty(values[16]) ? Convert.ToDouble(values[16]) : 0.0;
                            binData.HW1Max = !string.IsNullOrEmpty(values[17]) ? Convert.ToDouble(values[17]) : 0.0;
                            binData.LOP1Min = !string.IsNullOrEmpty(values[18]) ? Convert.ToDouble(values[18]) : 0.0;
                            binData.LOP1Max = !string.IsNullOrEmpty(values[19]) ? Convert.ToDouble(values[19]) : 0.0;
                            binData.WLP1Min = !string.IsNullOrEmpty(values[20]) ? Convert.ToDouble(values[20]) : 0.0;
                            binData.WLP1Max = !string.IsNullOrEmpty(values[21]) ? Convert.ToDouble(values[21]) : 0.0;
                            binData.WLD1Min = !string.IsNullOrEmpty(values[22]) ? Convert.ToDouble(values[22]) : 0.0;
                            binData.WLD1Max = !string.IsNullOrEmpty(values[23]) ? Convert.ToDouble(values[23]) : 0.0;
                            binData.IR1Min = !string.IsNullOrEmpty(values[24]) ? Convert.ToDouble(values[24]) : 0.0;
                            binData.IR1Max = !string.IsNullOrEmpty(values[25]) ? Convert.ToDouble(values[25]) : 0.0;
                            binData.VFDMin = !string.IsNullOrEmpty(values[26]) ? Convert.ToDouble(values[26]) : 0.0;
                            binData.VFDMax = !string.IsNullOrEmpty(values[27]) ? Convert.ToDouble(values[27]) : 0.0;
                            binData.DVFMin = !string.IsNullOrEmpty(values[28]) ? Convert.ToDouble(values[28]) : 0.0;
                            binData.DVFMax = !string.IsNullOrEmpty(values[29]) ? Convert.ToDouble(values[29]) : 0.0;
                            binData.IR2Min = !string.IsNullOrEmpty(values[30]) ? Convert.ToDouble(values[30]) : 0.0;
                            binData.IR2Max = !string.IsNullOrEmpty(values[31]) ? Convert.ToDouble(values[31]) : 0.0;
                            binData.WLC1Min = !string.IsNullOrEmpty(values[32]) ? Convert.ToDouble(values[32]) : 0.0;
                            binData.WLC1Max = !string.IsNullOrEmpty(values[33]) ? Convert.ToDouble(values[33]) : 0.0;
                            binData.VF5Min = !string.IsNullOrEmpty(values[34]) ? Convert.ToDouble(values[34]) : 0.0;
                            binData.VF5Max = !string.IsNullOrEmpty(values[35]) ? Convert.ToDouble(values[35]) : 0.0;
                            binData.VF6Min = !string.IsNullOrEmpty(values[36]) ? Convert.ToDouble(values[36]) : 0.0;
                            binData.VF6Max = !string.IsNullOrEmpty(values[37]) ? Convert.ToDouble(values[37]) : 0.0;
                            binData.VF7Min = !string.IsNullOrEmpty(values[38]) ? Convert.ToDouble(values[38]) : 0.0;
                            binData.VF7Max = !string.IsNullOrEmpty(values[39]) ? Convert.ToDouble(values[39]) : 0.0;
                            binData.VF8Min = !string.IsNullOrEmpty(values[40]) ? Convert.ToDouble(values[40]) : 0.0;
                            binData.VF8Max = !string.IsNullOrEmpty(values[41]) ? Convert.ToDouble(values[41]) : 0.0;
                            binData.DVF1Min = !string.IsNullOrEmpty(values[42]) ? Convert.ToDouble(values[42]) : 0.0;
                            binData.DVF1Max = !string.IsNullOrEmpty(values[43]) ? Convert.ToDouble(values[43]) : 0.0;
                            binData.DVF2Min = !string.IsNullOrEmpty(values[44]) ? Convert.ToDouble(values[44]) : 0.0;
                            binData.DVF2Max = !string.IsNullOrEmpty(values[45]) ? Convert.ToDouble(values[45]) : 0.0;
                            binData.VZ2Min = !string.IsNullOrEmpty(values[46]) ? Convert.ToDouble(values[46]) : 0.0;
                            binData.VZ2Max = !string.IsNullOrEmpty(values[47]) ? Convert.ToDouble(values[47]) : 0.0;
                            binData.VZ3Min = !string.IsNullOrEmpty(values[48]) ? Convert.ToDouble(values[48]) : 0.0;
                            binData.VZ3Max = !string.IsNullOrEmpty(values[49]) ? Convert.ToDouble(values[49]) : 0.0;
                            binData.VZ4Min = !string.IsNullOrEmpty(values[50]) ? Convert.ToDouble(values[50]) : 0.0;
                            binData.VZ4Max = !string.IsNullOrEmpty(values[51]) ? Convert.ToDouble(values[51]) : 0.0;
                            binData.VZ5Min = !string.IsNullOrEmpty(values[52]) ? Convert.ToDouble(values[52]) : 0.0;
                            binData.VZ5Max = !string.IsNullOrEmpty(values[53]) ? Convert.ToDouble(values[53]) : 0.0;
                            binData.IR3Min = !string.IsNullOrEmpty(values[54]) ? Convert.ToDouble(values[54]) : 0.0;
                            binData.IR3Max = !string.IsNullOrEmpty(values[55]) ? Convert.ToDouble(values[55]) : 0.0;
                            binData.IR4Min = !string.IsNullOrEmpty(values[56]) ? Convert.ToDouble(values[56]) : 0.0;
                            binData.IR4Max = !string.IsNullOrEmpty(values[57]) ? Convert.ToDouble(values[57]) : 0.0;
                            binData.IR5Min = !string.IsNullOrEmpty(values[58]) ? Convert.ToDouble(values[58]) : 0.0;
                            binData.IR5Max = !string.IsNullOrEmpty(values[59]) ? Convert.ToDouble(values[59]) : 0.0;
                            binData.IR6Min = !string.IsNullOrEmpty(values[60]) ? Convert.ToDouble(values[60]) : 0.0;
                            binData.IR6Max = !string.IsNullOrEmpty(values[61]) ? Convert.ToDouble(values[61]) : 0.0;
                            binData.IFMin = !string.IsNullOrEmpty(values[62]) ? Convert.ToDouble(values[62]) : 0.0;
                            binData.IFMax = !string.IsNullOrEmpty(values[63]) ? Convert.ToDouble(values[63]) : 0.0;
                            binData.IF1Min = !string.IsNullOrEmpty(values[64]) ? Convert.ToDouble(values[64]) : 0.0;
                            binData.IF1Max = !string.IsNullOrEmpty(values[65]) ? Convert.ToDouble(values[65]) : 0.0;
                            binData.IF2Min = !string.IsNullOrEmpty(values[66]) ? Convert.ToDouble(values[66]) : 0.0;
                            binData.IF2Max = !string.IsNullOrEmpty(values[67]) ? Convert.ToDouble(values[67]) : 0.0;
                            binData.LOP2Min = !string.IsNullOrEmpty(values[68]) ? Convert.ToDouble(values[68]) : 0.0;
                            binData.LOP2Max = !string.IsNullOrEmpty(values[69]) ? Convert.ToDouble(values[69]) : 0.0;
                            binData.WLP2Min = !string.IsNullOrEmpty(values[70]) ? Convert.ToDouble(values[70]) : 0.0;
                            binData.WLP2Max = !string.IsNullOrEmpty(values[71]) ? Convert.ToDouble(values[71]) : 0.0;
                            binData.WLD2Min = !string.IsNullOrEmpty(values[72]) ? Convert.ToDouble(values[72]) : 0.0;
                            binData.WLD2Max = !string.IsNullOrEmpty(values[73]) ? Convert.ToDouble(values[73]) : 0.0;
                            binData.HW2Min = !string.IsNullOrEmpty(values[74]) ? Convert.ToDouble(values[74]) : 0.0;
                            binData.HW2Max = !string.IsNullOrEmpty(values[75]) ? Convert.ToDouble(values[75]) : 0.0;
                            binData.WLC2Min = !string.IsNullOrEmpty(values[76]) ? Convert.ToDouble(values[76]) : 0.0;
                            binData.WLC2Max = !string.IsNullOrEmpty(values[77]) ? Convert.ToDouble(values[77]) : 0.0;

                            // 将bin表导入到list中
                            binDataList.Add(binData);
                        }
                    }
                    getMaxMin();

                }
                // 在循环结束后设置 ItemsSource
                binDataListBox.ItemsSource = binDataList;

                // 现在您可以在 binDataList 中访问导入的数据

                // 将最小值和最大值显示在 TextBox 中
                string data =  $"VF1Min: {vf1Min}, VF1Max: {vf1Max}\n" +
                        $"VF2Min: {vf2Min}, VF2Max: {vf2Max}\n" +
                        $"VF3Min: {vf3Min}, VF3Max: {vf3Max}\n" +
                        $"VF4Min: {vf4Min}, VF4Max: {vf4Max}\n" +
                        $"VZ1Min: {vz1Min}, VZ1Max: {vz1Max}\n" +
                        $"IRMin: {irMin}, IRMax: {irMax}\n" +
                        $"HW1Min: {hw1Min}, HW1Max: {hw1Max}\n" +
                        $"LOP1Min: {lop1Min}, LOP1Max: {lop1Max}\n" +
                        $"WLP1Min: {wlp1Min}, WLP1Max: {wlp1Max}\n" +
                        $"WLD1Min: {wld1Min}, WLD1Max: {wld1Max}\n" +
                        $"IR1Min: {ir1Min}, IR1Max: {ir1Max}\n" +
                        $"VFDMin: {vfdMin}, VFDMax: {vfdMax}\n" +
                        $"DVFMin: {dvfMin}, DVFMax: {dvfMax}\n" +
                        
                        $"IR2Min: {ir2Min}, IR2Max: {ir2Max}\n" +
                        $"WLC1Min: {wlc1Min}, WLC1Max: {wlc1Max}\n" +
                        $"VF5Min: {vf5Min}, VF5Max: {vf5Max}\n" +

                        $"VF6Min: {vf6Min}, VF6Max: {vf6Max}\n" +
                        $"VF7Min: {vf7Min}, VF7Max: {vf7Max}\n" +
                        $"VF8Min: {vf8Min}, VF8Max: {vf8Max}\n" +
                        $"DVF1Min: {dvf1Min}, DVF1Max: {dvf1Max}\n" +
                        $"DVF2Min: {dvf2Min}, DVF2Max: {dvf2Max}\n" +
                        $"VZ2Min: {vz2Min}, VZ2Max: {vz2Max}\n" +
                        $"VZ3Min: {vz3Min}, VZ3Max: {vz3Max}\n" +
                        $"VZ4Min: {vz4Min}, VZ4Max: {vz4Max}\n" +
                        $"VZ5Min: {vz5Min}, VZ5Max: {vz5Max}\n" +
                        $"IR3Min: {ir3Min}, IR3Max: {ir3Max}\n" +
                        $"IR4Min: {ir4Min}, IR4Max: {ir4Max}\n" +
                        $"IR5Min: {ir5Min}, IR5Max: {ir5Max}\n" +
                        $"IR6Min: {ir6Min}, IR6Max: {ir6Max}\n" +
                        $"IFMin: {ifMin}, IFMax: {ifMax}\n" +
                        $"IF1Min: {if1Min}, IF1Max: {if1Max}\n" +
                        $"IF2Min: {if2Min}, IF2Max: {if2Max}\n" +
                        $"LOP2Min: {lop2Min}, LOP2Max: {lop2Max}\n" +
                        $"WLP2Min: {wlp2Min}, WLP2Max: {wlp2Max}\n" +
                        $"WLD2Min: {wld2Min}, WLD2Max: {wld2Max}\n" +
                        $"HW2Min: {hw2Min}, HW2Max: {hw2Max}\n" +
                        $"WLC2Min: {wlc2Min}, WLC2Max: {wlc2Max}";

                string[] lines = data.Split('\n');
                foreach (string line in lines)
                {
                    string[] parts = line.Trim().Split(','); // 分割每一行，去掉首尾空格
                    if (parts.Length == 2)
                    {
                        parameterListBox.Items.Add(new DataItem { Min = parts[0], Max = parts[1] });
                    }
                }
                MessageBox.Show("Bin表文件导入成功！");
            }
            else
            {
                MessageBox.Show("请输入文件！");
            }
            
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            // 设置 LicenseContext 为 NonCommercial
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // 创建一个新的 ExcelPackage
            ExcelPackage excelPackage = new ExcelPackage();

            // 添加一个工作表
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

            // 在工作表中写入数据
            worksheet.Cells["A1"].Value = "Hello";
            worksheet.Cells["B1"].Value = "World";

            // 保存 ExcelPackage 到文件
            FileInfo excelFile = new FileInfo(@"C:\Users\Administrator\Desktop\落Bin软件\file.xlsx");
            excelPackage.SaveAs(excelFile);

            MessageBox.Show("Excel 文件已导出到 " + excelFile.FullName);
        }
        private void LoadFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";


            int totalChipNum = 0;
            if (openFileDialog.ShowDialog() == true)
            {
                List<Chip> chipList = new List<Chip>();

                foreach (string filename in openFileDialog.FileNames)
                {
                    using (StreamReader reader = new StreamReader(filename, Encoding.UTF8))
                    {
                        // 跳过前15行
                        for (int i = 0; i < 15; i++)
                        {
                            reader.ReadLine();
                        }

                        while (!reader.EndOfStream)
                        {
                            totalChipNum++;
                            string[] values = reader.ReadLine().Split(',');

                            // 创建一个新的 Chip 实例并设置属性值
                            Chip BinChip = new Chip();

                            BinChip.TEST = !string.IsNullOrEmpty(values[0]) ? Convert.ToDouble(values[0]) : 0.0;
                            BinChip.BIN = !string.IsNullOrEmpty(values[1]) ? Convert.ToDouble(values[1]) : 0.0;
                            BinChip.VF1 = !string.IsNullOrEmpty(values[2]) ? Convert.ToDouble(values[2]) : 0.0;
                            BinChip.VF2 = !string.IsNullOrEmpty(values[3]) ? Convert.ToDouble(values[3]) : 0.0;
                            BinChip.VF3 = !string.IsNullOrEmpty(values[4]) ? Convert.ToDouble(values[4]) : 0.0;
                            BinChip.VF4 = !string.IsNullOrEmpty(values[5]) ? Convert.ToDouble(values[5]) : 0.0;
                            BinChip.VF5 = !string.IsNullOrEmpty(values[6]) ? Convert.ToDouble(values[6]) : 0.0;
                            BinChip.VF6 = !string.IsNullOrEmpty(values[7]) ? Convert.ToDouble(values[7]) : 0.0;
                            BinChip.DVF = !string.IsNullOrEmpty(values[8]) ? Convert.ToDouble(values[8]) : 0.0;
                            BinChip.VF = !string.IsNullOrEmpty(values[9]) ? Convert.ToDouble(values[9]) : 0.0;
                            BinChip.VFD = !string.IsNullOrEmpty(values[10]) ? Convert.ToDouble(values[10]) : 0.0;
                            BinChip.VZ1 = !string.IsNullOrEmpty(values[11]) ? Convert.ToDouble(values[11]) : 0.0;
                            BinChip.VZ2 = !string.IsNullOrEmpty(values[12]) ? Convert.ToDouble(values[12]) : 0.0;
                            BinChip.IR = !string.IsNullOrEmpty(values[13]) ? Convert.ToDouble(values[13]) : 0.0;
                            BinChip.LOP1 = !string.IsNullOrEmpty(values[14]) ? Convert.ToDouble(values[14]) : 0.0;
                            BinChip.LOP2 = !string.IsNullOrEmpty(values[15]) ? Convert.ToDouble(values[15]) : 0.0;
                            BinChip.LOP3 = !string.IsNullOrEmpty(values[16]) ? Convert.ToDouble(values[16]) : 0.0;
                            BinChip.WLP1 = !string.IsNullOrEmpty(values[17]) ? Convert.ToDouble(values[17]) : 0.0;
                            BinChip.WLD1 = !string.IsNullOrEmpty(values[18]) ? Convert.ToDouble(values[18]) : 0.0;
                            BinChip.WLC1 = !string.IsNullOrEmpty(values[19]) ? Convert.ToDouble(values[19]) : 0.0;
                            BinChip.HW1 = !string.IsNullOrEmpty(values[20]) ? Convert.ToDouble(values[20]) : 0.0;
                            BinChip.WLP2 = !string.IsNullOrEmpty(values[27]) ? Convert.ToDouble(values[27]) : 0.0;
                            BinChip.WLD2 = !string.IsNullOrEmpty(values[28]) ? Convert.ToDouble(values[28]) : 0.0;
                            BinChip.WLC2 = !string.IsNullOrEmpty(values[29]) ? Convert.ToDouble(values[29]) : 0.0;
                            BinChip.HW2 = !string.IsNullOrEmpty(values[30]) ? Convert.ToDouble(values[30]) : 0.0;
                            BinChip.DVF1 = !string.IsNullOrEmpty(values[32]) ? Convert.ToDouble(values[32]) : 0.0;
                            BinChip.DVF2 = !string.IsNullOrEmpty(values[33]) ? Convert.ToDouble(values[33]) : 0.0;
                            BinChip.VF7 = !string.IsNullOrEmpty(values[36]) ? Convert.ToDouble(values[36]) : 0.0;
                            BinChip.VF8 = !string.IsNullOrEmpty(values[37]) ? Convert.ToDouble(values[37]) : 0.0;
                            BinChip.IR3 = !string.IsNullOrEmpty(values[38]) ? Convert.ToDouble(values[38]) : 0.0;
                            BinChip.IR4 = !string.IsNullOrEmpty(values[39]) ? Convert.ToDouble(values[39]) : 0.0;
                            BinChip.IR5 = !string.IsNullOrEmpty(values[40]) ? Convert.ToDouble(values[40]) : 0.0;
                            BinChip.IR6 = !string.IsNullOrEmpty(values[41]) ? Convert.ToDouble(values[41]) : 0.0;
                            BinChip.VZ3 = !string.IsNullOrEmpty(values[42]) ? Convert.ToDouble(values[42]) : 0.0;
                            BinChip.VZ4 = !string.IsNullOrEmpty(values[43]) ? Convert.ToDouble(values[43]) : 0.0;
                            BinChip.VZ5 = !string.IsNullOrEmpty(values[44]) ? Convert.ToDouble(values[44]) : 0.0;
                            BinChip.IF = !string.IsNullOrEmpty(values[45]) ? Convert.ToDouble(values[45]) : 0.0;
                            BinChip.IF1 = !string.IsNullOrEmpty(values[46]) ? Convert.ToDouble(values[46]) : 0.0;
                            BinChip.IF2 = !string.IsNullOrEmpty(values[47]) ? Convert.ToDouble(values[47]) : 0.0;
                            BinChip.IR1 = !string.IsNullOrEmpty(values[50]) ? Convert.ToDouble(values[50]) : 0.0;
                            BinChip.IR2 = !string.IsNullOrEmpty(values[51]) ? Convert.ToDouble(values[51]) : 0.0;

                            // 将 Chip 实例添加到列表中
                            chipList.Add(BinChip);
                        }
                    }
                }

                // 现在，chipList 包含所有 CSV 文件中的所有 Chip 实例
                // 您可以在这里执行其他操作，比如将它们显示在 ListBox 中
                foreach (Chip chip in chipList)
                {
                    parameterListBox.Items.Add(chip); 
                }
                MessageBox.Show("片号文件导入成功！");
            }
            else
            {
                MessageBox.Show("请输入文件！");
            }


        }
    }
}