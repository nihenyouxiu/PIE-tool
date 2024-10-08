﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using System.Windows.Controls;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.Intrinsics.X86;
using System.Reflection.Metadata;
using OfficeOpenXml.Style;
using OfficeOpenXml.DataValidation;
using System.Windows.Shapes;
using System.Drawing.Drawing2D;
using System.Text.RegularExpressions;
using System.Threading; // 添加多线程支持

namespace 产出分布计算
{
    /// <summary>
    /// Page5.xaml 的交互逻辑
    /// </summary>

    public partial class Page5 : Page
    {
        private static Dictionary<string, int> fieldOrder = new Dictionary<string, int>
        {
            {"TEST", 0}, {"BIN", 1}, {"VF1", 2}, {"VF2", 3}, {"VF3", 4}, {"VF4", 5}, {"VF5", 6}, {"VF6", 7}, {"DVF", 8},
            {"VF", 9}, {"VFD", 10}, {"VZ1", 11}, {"VZ2", 12}, {"IR", 13}, {"LOP1", 14}, {"LOP2", 15}, {"LOP3", 16},
            {"WLP1", 17}, {"WLD1", 18}, {"WLC1", 19}, {"HW1", 20}, {"PURITY1", 21}, {"X1", 22}, {"Y1", 23}, {"Z1", 24},
            {"ST1", 25}, {"INT1", 26}, {"WLP2", 27}, {"WLD2", 28}, {"WLC2", 29}, {"HW2", 30}, {"PURITY2", 31}, {"DVF1", 32},
            {"DVF2", 33}, {"INT2", 34}, {"ST2", 35}, {"VF7", 36}, {"VF8", 37}, {"IR3", 38}, {"IR4", 39}, {"IR5", 40}, {"IR6", 41},
            {"VZ3", 42}, {"VZ4", 43}, {"VZ5", 44}, {"IF", 45}, {"IF1", 46}, {"IF2", 47}, {"ESD1", 48}, {"ESD2", 49}, {"IR1", 50},
            {"IR2", 51}, {"ESD1PASS", 52}, {"ESD2PASS", 53}, {"PosX", 54}, {"PosY", 55}
        };
        List<BinData> binDataList;
        
        private readonly object binDataLock = new object(); // 添加锁对象用于保护binDataList
        private readonly object lockObject = new object();
        private readonly object parameterlockObject = new object();

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
        string postfix;

        public Page5()
        {
            InitializeComponent();
            this.KeepAlive = true;
        }

        public bool ValidateAgainstBinData(Chip chip, BinData binData)
        {
            return
            ((chip.VF1 == -1000000) || (binData.VF1Min == binData.VF1Max) || (chip.VF1 >= binData.VF1Min && chip.VF1 < binData.VF1Max)) &&
            ((chip.VF2 == -1000000) || (binData.VF2Min == binData.VF2Max) || (chip.VF2 >= binData.VF2Min && chip.VF2 < binData.VF2Max)) &&
            ((chip.VF3 == -1000000) || (binData.VF3Min == binData.VF3Max) || (chip.VF3 >= binData.VF3Min && chip.VF3 < binData.VF3Max)) &&
            ((chip.VF4 == -1000000) || (binData.VF4Min == binData.VF4Max) || (chip.VF4 >= binData.VF4Min && chip.VF4 < binData.VF4Max)) &&
            ((chip.VF5 == -1000000) || (binData.VF5Min == binData.VF5Max) || (chip.VF5 >= binData.VF5Min && chip.VF5 < binData.VF5Max)) &&
            ((chip.VF6 == -1000000) || (binData.VF6Min == binData.VF6Max) || (chip.VF6 >= binData.VF6Min && chip.VF6 < binData.VF6Max)) &&
            ((chip.DVF == -1000000) || (binData.DVFMin == binData.DVFMax) || (chip.DVF >= binData.DVFMin && chip.DVF < binData.DVFMax)) &&
            ((chip.VFD == -1000000) || (binData.VFDMin == binData.VFDMax) || (chip.VFD >= binData.VFDMin && chip.VFD < binData.VFDMax)) &&
            ((chip.VZ1 == -1000000) || (binData.VZ1Min == binData.VZ1Max) || (chip.VZ1 >= binData.VZ1Min && chip.VZ1 < binData.VZ1Max)) &&
            ((chip.VZ2 == -1000000) || (binData.VZ2Min == binData.VZ2Max) || (chip.VZ2 >= binData.VZ2Min && chip.VZ2 < binData.VZ2Max)) &&
            ((chip.IR == -1000000) || (binData.IRMin == binData.IRMax) || (chip.IR >= binData.IRMin && chip.IR < binData.IRMax)) &&
            ((chip.LOP1 == -1000000) || (binData.LOP1Min == binData.LOP1Max) || (chip.LOP1 >= binData.LOP1Min && chip.LOP1 < binData.LOP1Max)) &&
            ((chip.LOP2 == -1000000) || (binData.LOP2Min == binData.LOP2Max) || (chip.LOP2 >= binData.LOP2Min && chip.LOP2 < binData.LOP2Max)) &&
            ((chip.WLP1 == -1000000) || (binData.WLP1Min == binData.WLP1Max) || (chip.WLP1 >= binData.WLP1Min && chip.WLP1 < binData.WLP1Max)) &&
            ((chip.WLD1 == -1000000) || (binData.WLD1Min == binData.WLD1Max) || (chip.WLD1 >= binData.WLD1Min && chip.WLD1 < binData.WLD1Max)) &&
            ((chip.WLC1 == -1000000) || (binData.WLC1Min == binData.WLC1Max) || (chip.WLC1 >= binData.WLC1Min && chip.WLC1 < binData.WLC1Max)) &&
            ((chip.HW1 == -1000000) || (binData.HW1Min == binData.HW1Max) || (chip.HW1 >= binData.HW1Min && chip.HW1 < binData.HW1Max)) &&
            ((chip.WLP2 == -1000000) || (binData.WLP2Min == binData.WLP2Max) || (chip.WLP2 >= binData.WLP2Min && chip.WLP2 < binData.WLP2Max)) &&
            ((chip.WLD2 == -1000000) || (binData.WLD2Min == binData.WLD2Max) || (chip.WLD2 >= binData.WLD2Min && chip.WLD2 < binData.WLD2Max)) &&
            ((chip.WLC2 == -1000000) || (binData.WLC2Min == binData.WLC2Max) || (chip.WLC2 >= binData.WLC2Min && chip.WLC2 < binData.WLC2Max)) &&
            ((chip.HW2 == -1000000) || (binData.HW2Min == binData.HW2Max) || (chip.HW2 >= binData.HW2Min && chip.HW2 < binData.HW2Max)) &&
            ((chip.DVF1 == -1000000) || (binData.DVF1Min == binData.DVF1Max) || (chip.DVF1 >= binData.DVF1Min && chip.DVF1 < binData.DVF1Max)) &&
            ((chip.DVF2 == -1000000) || (binData.DVF2Min == binData.DVF2Max) || (chip.DVF2 >= binData.DVF2Min && chip.DVF2 < binData.DVF2Max)) &&
            ((chip.VF7 == -1000000) || (binData.VF7Min == binData.VF7Max) || (chip.VF7 >= binData.VF7Min && chip.VF7 < binData.VF7Max)) &&
            ((chip.VF8 == -1000000) || (binData.VF8Min == binData.VF8Max) || (chip.VF8 >= binData.VF8Min && chip.VF8 < binData.VF8Max)) &&
            ((chip.IR3 == -1000000) || (binData.IR3Min == binData.IR3Max) || (chip.IR3 >= binData.IR3Min && chip.IR3 < binData.IR3Max)) &&
            ((chip.IR4 == -1000000) || (binData.IR4Min == binData.IR4Max) || (chip.IR4 >= binData.IR4Min && chip.IR4 < binData.IR4Max)) &&
            ((chip.IR5 == -1000000) || (binData.IR5Min == binData.IR5Max) || (chip.IR5 >= binData.IR5Min && chip.IR5 < binData.IR5Max)) &&
            ((chip.IR6 == -1000000) || (binData.IR6Min == binData.IR6Max) || (chip.IR6 >= binData.IR6Min && chip.IR6 < binData.IR6Max)) &&
            ((chip.VZ3 == -1000000) || (binData.VZ3Min == binData.VZ3Max) || (chip.VZ3 >= binData.VZ3Min && chip.VZ3 < binData.VZ3Max)) &&
            ((chip.VZ4 == -1000000) || (binData.VZ4Min == binData.VZ4Max) || (chip.VZ4 >= binData.VZ4Min && chip.VZ4 < binData.VZ4Max)) &&
            ((chip.VZ5 == -1000000) || (binData.VZ5Min == binData.VZ5Max) || (chip.VZ5 >= binData.VZ5Min && chip.VZ5 < binData.VZ5Max)) &&
            ((chip.IF == -1000000) || (binData.IFMin == binData.IFMax) || (chip.IF >= binData.IFMin && chip.IF < binData.IFMax)) &&
            ((chip.IF1 == -1000000) || (binData.IF1Min == binData.IF1Max) || (chip.IF1 >= binData.IF1Min && chip.IF1 < binData.IF1Max)) &&
            ((chip.IF2 == -1000000) || (binData.IF2Min == binData.IF2Max) || (chip.IF2 >= binData.IF2Min && chip.IF2 < binData.IF2Max)) &&
            ((chip.IR1 == -1000000) || (binData.IR1Min == binData.IR1Max) || (chip.IR1 >= binData.IR1Min && chip.IR1 < binData.IR1Max)) &&
            ((chip.IR2 == -1000000) || (binData.IR2Min == binData.IR2Max) || (chip.IR2 >= binData.IR2Min && chip.IR2 < binData.IR2Max)) && true;
        }
        void getMaxMin()
        {
            if (binDataList.Any())
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
            else
            {
                MessageBox.Show("输入文件有误，请重试！");
                return;
            }
        }

        string chip_ToString(Chip chip)
        {

            string formattedString = $",{chip.TEST},{postfix}{chip.BIN:000},{chip.VF1},{chip.VF2},{chip.VF3},{chip.VF4},{chip.VF5},{chip.VF6},{chip.DVF},{chip.VF},{chip.VFD},{chip.VZ1},{chip.VZ2},{chip.IR},{chip.LOP1},{chip.LOP2},{chip.LOP3},{chip.WLP1},{chip.WLD1},{chip.WLC1},{chip.HW1},{chip.WLP2},{chip.WLD2},{chip.WLC2},{chip.HW2},{chip.DVF1},{chip.DVF2},{chip.VF7},{chip.VF8},{chip.IR3},{chip.IR4},{chip.IR5},{chip.IR6},{chip.VZ3},{chip.VZ4},{chip.VZ5},{chip.IF},{chip.IF1},{chip.IF2},{chip.IR1},{chip.IR2}";

            return formattedString;
        }
        string output_excel_file_name;
        private void BinImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "CSV files (*.csv)|*.csv";
            if (openFileDialog.ShowDialog() == true)
            {
                string filename = openFileDialog.FileName;
                binDataList = new List<BinData>();
                try
                {
                    using (var reader = new StreamReader(filename, Encoding.UTF8))
                    {
                        string output_csv_file_name = System.IO.Path.GetFileNameWithoutExtension(filename); // 获取文件名，不含扩展名
                        postfix = output_csv_file_name.Substring(output_csv_file_name.Length - 2);
                        output_excel_file_name = $"{output_csv_file_name}.xlsx"; // 添加新的后缀名
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
                            if (line.StartsWith("1") && values.Length >= 78)
                            {
                                BinData binData = new BinData();
                                //将每个字段的值分配给相应的属性
                                binData.binIdx = !string.IsNullOrEmpty(values[2]) ? Convert.ToDouble(values[2]) : -100000;
                                binData.VF1Min = !string.IsNullOrEmpty(values[4]) ? Convert.ToDouble(values[4]) : -100000;
                                binData.VF1Max = !string.IsNullOrEmpty(values[5]) ? Convert.ToDouble(values[5]) : -100000;
                                binData.VF2Min = !string.IsNullOrEmpty(values[6]) ? Convert.ToDouble(values[6]) : -100000;
                                binData.VF2Max = !string.IsNullOrEmpty(values[7]) ? Convert.ToDouble(values[7]) : -100000;
                                binData.VF3Min = !string.IsNullOrEmpty(values[8]) ? Convert.ToDouble(values[8]) : -100000;
                                binData.VF3Max = !string.IsNullOrEmpty(values[9]) ? Convert.ToDouble(values[9]) : -100000;
                                binData.VF4Min = !string.IsNullOrEmpty(values[10]) ? Convert.ToDouble(values[10]) : -100000;
                                binData.VF4Max = !string.IsNullOrEmpty(values[11]) ? Convert.ToDouble(values[11]) : -100000;
                                binData.VZ1Min = !string.IsNullOrEmpty(values[12]) ? Convert.ToDouble(values[12]) : -100000;
                                binData.VZ1Max = !string.IsNullOrEmpty(values[13]) ? Convert.ToDouble(values[13]) : -100000;
                                binData.IRMin = !string.IsNullOrEmpty(values[14]) ? Convert.ToDouble(values[14]) : -100000;
                                binData.IRMax = !string.IsNullOrEmpty(values[15]) ? Convert.ToDouble(values[15]) : -100000;
                                binData.HW1Min = !string.IsNullOrEmpty(values[16]) ? Convert.ToDouble(values[16]) : -100000;
                                binData.HW1Max = !string.IsNullOrEmpty(values[17]) ? Convert.ToDouble(values[17]) : -100000;
                                binData.LOP1Min = !string.IsNullOrEmpty(values[18]) ? Convert.ToDouble(values[18]) : -100000;
                                binData.LOP1Max = !string.IsNullOrEmpty(values[19]) ? Convert.ToDouble(values[19]) : -100000;
                                binData.WLP1Min = !string.IsNullOrEmpty(values[20]) ? Convert.ToDouble(values[20]) : -100000;
                                binData.WLP1Max = !string.IsNullOrEmpty(values[21]) ? Convert.ToDouble(values[21]) : -100000;
                                binData.WLD1Min = !string.IsNullOrEmpty(values[22]) ? Convert.ToDouble(values[22]) : -100000;
                                binData.WLD1Max = !string.IsNullOrEmpty(values[23]) ? Convert.ToDouble(values[23]) : -100000;
                                binData.IR1Min = !string.IsNullOrEmpty(values[24]) ? Convert.ToDouble(values[24]) : -100000;
                                binData.IR1Max = !string.IsNullOrEmpty(values[25]) ? Convert.ToDouble(values[25]) : -100000;
                                binData.VFDMin = !string.IsNullOrEmpty(values[26]) ? Convert.ToDouble(values[26]) : -100000;
                                binData.VFDMax = !string.IsNullOrEmpty(values[27]) ? Convert.ToDouble(values[27]) : -100000;
                                binData.DVFMin = !string.IsNullOrEmpty(values[28]) ? Convert.ToDouble(values[28]) : -100000;
                                binData.DVFMax = !string.IsNullOrEmpty(values[29]) ? Convert.ToDouble(values[29]) : -100000;
                                binData.IR2Min = !string.IsNullOrEmpty(values[30]) ? Convert.ToDouble(values[30]) : -100000;
                                binData.IR2Max = !string.IsNullOrEmpty(values[31]) ? Convert.ToDouble(values[31]) : -100000;
                                binData.WLC1Min = !string.IsNullOrEmpty(values[32]) ? Convert.ToDouble(values[32]) : -100000;
                                binData.WLC1Max = !string.IsNullOrEmpty(values[33]) ? Convert.ToDouble(values[33]) : -100000;
                                binData.VF5Min = !string.IsNullOrEmpty(values[34]) ? Convert.ToDouble(values[34]) : -100000;
                                binData.VF5Max = !string.IsNullOrEmpty(values[35]) ? Convert.ToDouble(values[35]) : -100000;
                                binData.VF6Min = !string.IsNullOrEmpty(values[36]) ? Convert.ToDouble(values[36]) : -100000;
                                binData.VF6Max = !string.IsNullOrEmpty(values[37]) ? Convert.ToDouble(values[37]) : -100000;
                                binData.VF7Min = !string.IsNullOrEmpty(values[38]) ? Convert.ToDouble(values[38]) : -100000;
                                binData.VF7Max = !string.IsNullOrEmpty(values[39]) ? Convert.ToDouble(values[39]) : -100000;
                                binData.VF8Min = !string.IsNullOrEmpty(values[40]) ? Convert.ToDouble(values[40]) : -100000;
                                binData.VF8Max = !string.IsNullOrEmpty(values[41]) ? Convert.ToDouble(values[41]) : -100000;
                                binData.DVF1Min = !string.IsNullOrEmpty(values[42]) ? Convert.ToDouble(values[42]) : -100000;
                                binData.DVF1Max = !string.IsNullOrEmpty(values[43]) ? Convert.ToDouble(values[43]) : -100000;
                                binData.DVF2Min = !string.IsNullOrEmpty(values[44]) ? Convert.ToDouble(values[44]) : -100000;
                                binData.DVF2Max = !string.IsNullOrEmpty(values[45]) ? Convert.ToDouble(values[45]) : -100000;
                                binData.VZ2Min = !string.IsNullOrEmpty(values[46]) ? Convert.ToDouble(values[46]) : -100000;
                                binData.VZ2Max = !string.IsNullOrEmpty(values[47]) ? Convert.ToDouble(values[47]) : -100000;
                                binData.VZ3Min = !string.IsNullOrEmpty(values[48]) ? Convert.ToDouble(values[48]) : -100000;
                                binData.VZ3Max = !string.IsNullOrEmpty(values[49]) ? Convert.ToDouble(values[49]) : -100000;
                                binData.VZ4Min = !string.IsNullOrEmpty(values[50]) ? Convert.ToDouble(values[50]) : -100000;
                                binData.VZ4Max = !string.IsNullOrEmpty(values[51]) ? Convert.ToDouble(values[51]) : -100000;
                                binData.VZ5Min = !string.IsNullOrEmpty(values[52]) ? Convert.ToDouble(values[52]) : -100000;
                                binData.VZ5Max = !string.IsNullOrEmpty(values[53]) ? Convert.ToDouble(values[53]) : -100000;
                                binData.IR3Min = !string.IsNullOrEmpty(values[54]) ? Convert.ToDouble(values[54]) : -100000;
                                binData.IR3Max = !string.IsNullOrEmpty(values[55]) ? Convert.ToDouble(values[55]) : -100000;
                                binData.IR4Min = !string.IsNullOrEmpty(values[56]) ? Convert.ToDouble(values[56]) : -100000;
                                binData.IR4Max = !string.IsNullOrEmpty(values[57]) ? Convert.ToDouble(values[57]) : -100000;
                                binData.IR5Min = !string.IsNullOrEmpty(values[58]) ? Convert.ToDouble(values[58]) : -100000;
                                binData.IR5Max = !string.IsNullOrEmpty(values[59]) ? Convert.ToDouble(values[59]) : -100000;
                                binData.IR6Min = !string.IsNullOrEmpty(values[60]) ? Convert.ToDouble(values[60]) : -100000;
                                binData.IR6Max = !string.IsNullOrEmpty(values[61]) ? Convert.ToDouble(values[61]) : -100000;
                                binData.IFMin = !string.IsNullOrEmpty(values[62]) ? Convert.ToDouble(values[62]) : -100000;
                                binData.IFMax = !string.IsNullOrEmpty(values[63]) ? Convert.ToDouble(values[63]) : -100000;
                                binData.IF1Min = !string.IsNullOrEmpty(values[64]) ? Convert.ToDouble(values[64]) : -100000;
                                binData.IF1Max = !string.IsNullOrEmpty(values[65]) ? Convert.ToDouble(values[65]) : -100000;
                                binData.IF2Min = !string.IsNullOrEmpty(values[66]) ? Convert.ToDouble(values[66]) : -100000;
                                binData.IF2Max = !string.IsNullOrEmpty(values[67]) ? Convert.ToDouble(values[67]) : -100000;
                                binData.LOP2Min = !string.IsNullOrEmpty(values[68]) ? Convert.ToDouble(values[68]) : -100000;
                                binData.LOP2Max = !string.IsNullOrEmpty(values[69]) ? Convert.ToDouble(values[69]) : -100000;
                                binData.WLP2Min = !string.IsNullOrEmpty(values[70]) ? Convert.ToDouble(values[70]) : -100000;
                                binData.WLP2Max = !string.IsNullOrEmpty(values[71]) ? Convert.ToDouble(values[71]) : -100000;
                                binData.WLD2Min = !string.IsNullOrEmpty(values[72]) ? Convert.ToDouble(values[72]) : -100000;
                                binData.WLD2Max = !string.IsNullOrEmpty(values[73]) ? Convert.ToDouble(values[73]) : -100000;
                                binData.HW2Min = !string.IsNullOrEmpty(values[74]) ? Convert.ToDouble(values[74]) : -100000;
                                binData.HW2Max = !string.IsNullOrEmpty(values[75]) ? Convert.ToDouble(values[75]) : -100000;
                                binData.WLC2Min = !string.IsNullOrEmpty(values[76]) ? Convert.ToDouble(values[76]) : -100000;
                                binData.WLC2Max = !string.IsNullOrEmpty(values[77]) ? Convert.ToDouble(values[77]) : -100000;
                                binData.chipNum = 0;

                                // 将bin表导入到list中
                                binDataList.Add(binData);
                            }
                        }
                        getMaxMin();
                        binDatafail.binIdx = 999;
                        binDataList.Add(binDatafail);

                    }
                }
                catch (IOException)
                {
                    MessageBox.Show($"文件 {filename} 已被打开，请关闭后重新选择!", "文件已打开", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }

                binDataListBox.ItemsSource = binDataList;
                MessageBox.Show("Bin表文件导入成功，请载入片号文件！");

            }
            else
            {
                MessageBox.Show("请输入文件！");
            }

        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            if (binDataList == null)
            {
                MessageBox.Show("导出文件失败，请输入文件！", "未输入文件");
                return;
            }
            string outputFolder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutputFolder");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage();
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

            worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            worksheet.Cells[1, 1].Value = "BIN";
            worksheet.Cells[1, 2].Value = "WLD1";
            worksheet.Cells[1, 3].Value = "WLP1";
            worksheet.Cells[1, 4].Value = "LOP1";
            worksheet.Cells[1, 5].Value = "VF1";
            worksheet.Cells[1, 6].Value = "VF2";
            worksheet.Cells[1, 7].Value = "VF3";
            worksheet.Cells[1, 8].Value = "ChipNum";
            worksheet.Cells[1, 9].Value = "落bin率";

            for (int col = 1; col <= 9; col++)
            {
                worksheet.Cells[1, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                worksheet.Cells[1, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            }

            // 设置第一行的字体为微软雅黑、大小为14号
            using (ExcelRange range = worksheet.Cells["A1:I1"])
            {
                range.Style.Font.Name = "微软雅黑";
                range.Style.Font.Size = 14;
            }

            // 设置第一行的填充颜色为浅蓝色
            using (ExcelRange range = worksheet.Cells["A1:I1"])
            {
                var fill = range.Style.Fill;
                fill.PatternType = ExcelFillStyle.Solid;
                fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 204, 255));
            }

            // 写入属性值到第二行开始
            int row = 2;
            foreach (var binData in binDataList)
            {
                worksheet.Cells[row, 1].Value = binData.binIdx;
                worksheet.Cells[row, 2].Value = $"[{binData.WLD1Min} , {binData.WLD1Max})";
                worksheet.Cells[row, 3].Value = $"[{binData.WLP1Min} , {binData.WLP1Max})";
                worksheet.Cells[row, 4].Value = $"[{binData.LOP1Min} , {binData.LOP1Max})";
                worksheet.Cells[row, 5].Value = $"[{binData.VF1Min} , {binData.VF1Max})";
                worksheet.Cells[row, 6].Value = $"[{binData.VF2Min} , {binData.VF2Max})";
                worksheet.Cells[row, 7].Value = $"[{binData.VF3Min} , {binData.VF3Max})";
                worksheet.Cells[row, 8].Value = binData.chipNum;
                worksheet.Cells[row, 9].Value = (double)binData.chipNum / totalChipNum;
                // 将第九列的格式更改为数字
                worksheet.Cells[row, 9].Style.Numberformat.Format = "0.00%";
                row++;
            }
            worksheet.Cells[row, 1].Value = "total";
            worksheet.Cells[row, 8].Value = totalChipNum - binDatafail.chipNum;

            worksheet.Cells[row, 9].Value = (totalChipNum - (double)binDatafail.chipNum) / totalChipNum;
            worksheet.Cells[row, 9].Style.Numberformat.Format = "0.00%";
            // 自动调整列宽以适应内容
            worksheet.Cells.AutoFitColumns();
            // 保存 ExcelPackage 到文件

            string output_excel_file = System.IO.Path.Combine(outputFolder, output_excel_file_name);

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
                // 弹出消息框询问是否打开文件
                MessageBoxResult result = MessageBox.Show("Excel 文件已导出到 " + output_excel_file + "\n是否打开该文件？", "导出成功", MessageBoxButton.YesNo, MessageBoxImage.Question);

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
            }
            else
            {
                // 处理文件名为空的情况
                MessageBox.Show("Excel 文件: " + output_excel_file + "导出出错！");
            }
        }


        BinData binDatafail = new BinData();
        double totalChipNum = 0;
        bool breakFlag = false;
        
        private async void ProcessFile(string filename, string outputCsvFile, double vf1fixNum, double lop1fixNum)
        {
            List<Chip> chipList = new List<Chip>();
            int minX = int.MaxValue;
            int maxX = int.MinValue;
            int minY = int.MaxValue;
            int maxY = int.MinValue;
            string lines = "";
            Dictionary<(int X, int Y), Chip> chipDictionary = new Dictionary<(int X, int Y), Chip>();
            int flag = 0;
            try
            {
                using (StreamReader reader = new StreamReader(filename))
                {

                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        string[] values = line.Split(',');


                        if (values.Length >= 56 && values[0] == "TEST")
                        {
                            if (values[1] == "BIN1" && values[2] == "BIN2")
                                flag = 1;
                        }
                        string firstValue = values[0];
                        bool isFirstValueAllDigits = Regex.IsMatch(firstValue, @"^\d+$");
                        if (!isFirstValueAllDigits)
                        {
                            lines += line + "\r\n";
                        }
                        if (isFirstValueAllDigits && values.Length >= 56)
                        {
                            Chip chipData = new Chip
                            {
                                TEST = !string.IsNullOrEmpty(values[0]) ? Convert.ToDouble(values[0]) : -100000,
                                BIN = !string.IsNullOrEmpty(values[1]) ? 999 : -100000,
                                VF1 = !string.IsNullOrEmpty(values[2 + flag]) ? Convert.ToDouble(values[2 + flag]) : -100000,
                                VF2 = !string.IsNullOrEmpty(values[3 + flag]) ? Convert.ToDouble(values[3 + flag]) : -100000,
                                VF3 = !string.IsNullOrEmpty(values[4 + flag]) ? Convert.ToDouble(values[4 + flag]) : -100000,
                                VF4 = !string.IsNullOrEmpty(values[5 + flag]) ? Convert.ToDouble(values[5 + flag]) : -100000,
                                VF5 = !string.IsNullOrEmpty(values[6 + flag]) ? Convert.ToDouble(values[6 + flag]) : -100000,
                                VF6 = !string.IsNullOrEmpty(values[7 + flag]) ? Convert.ToDouble(values[7 + flag]) : -100000,
                                DVF = !string.IsNullOrEmpty(values[8 + flag]) ? Convert.ToDouble(values[8 + flag]) : -100000,
                                VF = !string.IsNullOrEmpty(values[9 + flag]) ? Convert.ToDouble(values[9 + flag]) : -100000,
                                VFD = !string.IsNullOrEmpty(values[10 + flag]) ? Convert.ToDouble(values[10 + flag]) : -100000,
                                VZ1 = !string.IsNullOrEmpty(values[11 + flag]) ? Convert.ToDouble(values[11 + flag]) : -100000,
                                VZ2 = !string.IsNullOrEmpty(values[12 + flag]) ? Convert.ToDouble(values[12 + flag]) : -100000,
                                IR = !string.IsNullOrEmpty(values[13 + flag]) ? Convert.ToDouble(values[13 + flag]) : -100000,
                                LOP1 = !string.IsNullOrEmpty(values[14 + flag]) ? Convert.ToDouble(values[14 + flag]) : -100000,
                                LOP2 = !string.IsNullOrEmpty(values[15 + flag]) ? Convert.ToDouble(values[15 + flag]) : -100000,
                                LOP3 = !string.IsNullOrEmpty(values[16 + flag]) ? Convert.ToDouble(values[16 + flag]) : -100000,
                                WLP1 = !string.IsNullOrEmpty(values[17 + flag]) ? Convert.ToDouble(values[17 + flag]) : -100000,
                                WLD1 = !string.IsNullOrEmpty(values[18 + flag]) ? Convert.ToDouble(values[18 + flag]) : -100000,
                                WLC1 = !string.IsNullOrEmpty(values[19 + flag]) ? Convert.ToDouble(values[19 + flag]) : -100000,
                                HW1 = !string.IsNullOrEmpty(values[20 + flag]) ? Convert.ToDouble(values[20 + flag]) : -100000,
                                PURITY1 = !string.IsNullOrEmpty(values[21 + flag]) ? Convert.ToDouble(values[21 + flag]) : -100000,
                                X1 = !string.IsNullOrEmpty(values[22 + flag]) ? Convert.ToDouble(values[22 + flag]) : -100000,
                                Y1 = !string.IsNullOrEmpty(values[23 + flag]) ? Convert.ToDouble(values[23 + flag]) : -100000,
                                Z1 = !string.IsNullOrEmpty(values[24 + flag]) ? Convert.ToDouble(values[24 + flag]) : -100000,
                                ST1 = !string.IsNullOrEmpty(values[25 + flag]) ? Convert.ToDouble(values[25 + flag]) : -100000,
                                INT1 = !string.IsNullOrEmpty(values[26 + flag]) ? Convert.ToDouble(values[26 + flag]) : -100000,
                                WLP2 = !string.IsNullOrEmpty(values[27 + flag]) ? Convert.ToDouble(values[27 + flag]) : -100000,
                                WLD2 = !string.IsNullOrEmpty(values[28 + flag]) ? Convert.ToDouble(values[28 + flag]) : -100000,
                                WLC2 = !string.IsNullOrEmpty(values[29 + flag]) ? Convert.ToDouble(values[29 + flag]) : -100000,
                                HW2 = !string.IsNullOrEmpty(values[30 + flag]) ? Convert.ToDouble(values[30 + flag]) : -100000,
                                PURITY2 = !string.IsNullOrEmpty(values[31 + flag]) ? Convert.ToDouble(values[31 + flag]) : -100000,
                                DVF1 = !string.IsNullOrEmpty(values[32 + flag]) ? Convert.ToDouble(values[32 + flag]) : -100000,
                                DVF2 = !string.IsNullOrEmpty(values[33 + flag]) ? Convert.ToDouble(values[33 + flag]) : -100000,
                                INT2 = !string.IsNullOrEmpty(values[34 + flag]) ? Convert.ToDouble(values[34 + flag]) : -100000,
                                ST2 = !string.IsNullOrEmpty(values[35 + flag]) ? Convert.ToDouble(values[35 + flag]) : -100000,
                                VF7 = !string.IsNullOrEmpty(values[36 + flag]) ? Convert.ToDouble(values[36 + flag]) : -100000,
                                VF8 = !string.IsNullOrEmpty(values[37 + flag]) ? Convert.ToDouble(values[37 + flag]) : -100000,
                                IR3 = !string.IsNullOrEmpty(values[38 + flag]) ? Convert.ToDouble(values[38 + flag]) : -100000,
                                IR4 = !string.IsNullOrEmpty(values[39 + flag]) ? Convert.ToDouble(values[39 + flag]) : -100000,
                                IR5 = !string.IsNullOrEmpty(values[40 + flag]) ? Convert.ToDouble(values[40 + flag]) : -100000,
                                IR6 = !string.IsNullOrEmpty(values[41 + flag]) ? Convert.ToDouble(values[41 + flag]) : -100000,
                                VZ3 = !string.IsNullOrEmpty(values[42 + flag]) ? Convert.ToDouble(values[42 + flag]) : -100000,
                                VZ4 = !string.IsNullOrEmpty(values[43 + flag]) ? Convert.ToDouble(values[43 + flag]) : -100000,
                                VZ5 = !string.IsNullOrEmpty(values[44 + flag]) ? Convert.ToDouble(values[44 + flag]) : -100000,
                                IF = !string.IsNullOrEmpty(values[45 + flag]) ? Convert.ToDouble(values[45 + flag]) : -100000,
                                IF1 = !string.IsNullOrEmpty(values[46 + flag]) ? Convert.ToDouble(values[46 + flag]) : -100000,
                                IF2 = !string.IsNullOrEmpty(values[47 + flag]) ? Convert.ToDouble(values[47 + flag]) : -100000,
                                ESD1 = !string.IsNullOrEmpty(values[48 + flag]) ? Convert.ToDouble(values[48 + flag]) : -100000,
                                ESD2 = !string.IsNullOrEmpty(values[49 + flag]) ? Convert.ToDouble(values[49 + flag]) : -100000,
                                IR1 = !string.IsNullOrEmpty(values[50 + flag]) ? Convert.ToDouble(values[50 + flag]) : -100000,
                                IR2 = !string.IsNullOrEmpty(values[51 + flag]) ? Convert.ToDouble(values[51 + flag]) : -100000,
                                ESD1PASS = !string.IsNullOrEmpty(values[52 + flag]) ? Convert.ToDouble(values[52 + flag]) : -100000,
                                ESD2PASS = !string.IsNullOrEmpty(values[53 + flag]) ? Convert.ToDouble(values[53 + flag]) : -100000,
                                PosX = !string.IsNullOrEmpty(values[54 + flag]) ? Convert.ToInt32(values[54 + flag]) : -100000,
                                PosY = !string.IsNullOrEmpty(values[55 + flag]) ? Convert.ToInt32(values[55 + flag]) : -100000,
                            };
                            chipList.Add(chipData);

                            chipDictionary[(chipData.PosX, chipData.PosY)] = chipData;
                            
                            if (chipData.PosX < minX) minX = chipData.PosX; // Track min and max values of PosX and PosY
                            if (chipData.PosX > maxX) maxX = chipData.PosX;
                            if (chipData.PosY < minY) minY = chipData.PosY;
                            if (chipData.PosY > maxY) maxY = chipData.PosY;
                        }
                    }
                }
            }
            catch (IOException)
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    MessageBox.Show($"文件 {filename} 已被打开，请关闭后重新选择!", "文件已打开", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                });
                return;
            }

            // Ensure minX and minY are even
            minX = minX % 2 == 0 ? minX : minX - 1;
            minY = minY % 2 == 0 ? minY : minY - 1;

            // Ensure maxX and maxY are even
            maxX = maxX % 2 == 0 ? maxX : maxX + 1;
            maxY = maxY % 2 == 0 ? maxY : maxY + 1;

            BinData binDataFailTmp = binDataList.FirstOrDefault(item => item.binIdx == 999);
            int waferidchipnum = 0;
            if (chipList.Any())
            {
                foreach (Chip chip in chipList)
                {
                    bool flag1 = false;

                    foreach (BinData binDataTmp in binDataList)
                    {
                        if (ValidateAgainstBinData(chip, binDataTmp))
                        {
                            lock (binDataTmp.Lock)
                            {
                                binDataTmp.chipNum++;
                                chip.BIN = binDataTmp.binIdx;
                            }
                            flag1 = true;
                            break;
                        }
                    }

                    if (!flag1)
                    {
                        lock (binDataFailTmp?.Lock)
                        {
                            if (binDataFailTmp != null)
                            {
                                chip.BIN = 999;
                                binDataFailTmp.chipNum++;
                            }
                        }
                    }

                    lock (lockObject)
                    {
                        totalChipNum++;
                    }
                    waferidchipnum++;
                }
                chipList.Clear();
            }
            else
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    breakFlag = true;
                    MessageBox.Show("输入文件有误，请重新输入！");
                });
            }
            lines = lines.TrimEnd('\r', '\n');
            await Task.Run(() =>
            {
                lock (lockObject)
                {
                    using (StreamWriter sw = new StreamWriter(outputCsvFile, true, Encoding.UTF8))
                    {
                        sw.WriteLineAsync(lines);
                    }
                }
            });

            int index = 0;
            for (int x = minX; x <= maxX; x += 2)
            {
                for (int y = minY; y <= maxY; y += 2)
                {
                    if (chipDictionary.ContainsKey((x, y)) && chipDictionary.ContainsKey((x, y + 1)) &&
                        chipDictionary.ContainsKey((x + 1, y)) && chipDictionary.ContainsKey((x + 1, y + 1)))
                    {
                        index++;
                        var chips = new List<Chip>
                            {
                                chipDictionary[(x, y)],
                                chipDictionary[(x, y + 1)],
                                chipDictionary[(x + 1, y)],
                                chipDictionary[(x + 1, y + 1)]
                            };

                        var averagedChip = new Chip
                        {
                            TEST = index,
                            BIN = -100000,
                            VF1 = 0.0f,
                            VF2 = chips.Average(c => c.VF2),
                            VF3 = chips.Average(c => c.VF3),
                            VF4 = chips.Average(c => c.VF4),
                            VF5 = chips.Average(c => c.VF5),
                            VF6 = chips.Average(c => c.VF6),
                            DVF = chips.Average(c => c.DVF),
                            VF = chips.Average(c => c.VF),
                            VFD = chips.Average(c => c.VFD),
                            VZ1 = chips.Average(c => c.VZ1),
                            VZ2 = chips.Average(c => c.VZ2),
                            IR = chips.Average(c => c.IR),
                            LOP1 = chips.Average(c => c.LOP1),
                            LOP2 = chips.Average(c => c.LOP2),
                            LOP3 = chips.Average(c => c.LOP3),
                            WLP1 = chips.Average(c => c.WLP1),
                            WLD1 = chips.Average(c => c.WLD1),
                            WLC1 = chips.Average(c => c.WLC1),
                            HW1 = chips.Average(c => c.HW1),
                            PURITY1 = chips.Average(c => c.PURITY1),
                            X1 = chips.Average(c => c.X1),
                            Y1 = chips.Average(c => c.Y1),
                            Z1 = chips.Average(c => c.Z1),
                            ST1 = chips.Average(c => c.ST1),
                            INT1 = chips.Average(c => c.INT1),
                            WLP2 = chips.Average(c => c.WLP2),
                            WLD2 = chips.Average(c => c.WLD2),
                            WLC2 = chips.Average(c => c.WLC2),
                            HW2 = chips.Average(c => c.HW2),
                            PURITY2 = chips.Average(c => c.PURITY2),
                            DVF1 = chips.Average(c => c.DVF1),
                            DVF2 = chips.Average(c => c.DVF2),
                            INT2 = chips.Average(c => c.INT2),
                            ST2 = chips.Average(c => c.ST2),
                            VF7 = chips.Average(c => c.VF7),
                            VF8 = chips.Average(c => c.VF8),
                            IR3 = chips.Average(c => c.IR3),
                            IR4 = chips.Average(c => c.IR4),
                            IR5 = chips.Average(c => c.IR5),
                            IR6 = chips.Average(c => c.IR6),
                            VZ3 = chips.Average(c => c.VZ3),
                            VZ4 = chips.Average(c => c.VZ4),
                            VZ5 = chips.Average(c => c.VZ5),
                            IF = chips.Average(c => c.IF),
                            IF1 = chips.Average(c => c.IF1),
                            IF2 = chips.Average(c => c.IF2),
                            ESD1 = chips.Average(c => c.ESD1),
                            ESD2 = chips.Average(c => c.ESD2),
                            IR1 = chips.Average(c => c.IR1),
                            IR2 = chips.Average(c => c.IR2),
                            ESD1PASS = chips.Average(c => c.ESD1PASS),
                            ESD2PASS = chips.Average(c => c.ESD2PASS),
                            PosX = x / 2 ,
                            PosY = y / 2
                        };

                        if (chipDictionary[(x,y)].BIN != 999 && chipDictionary[(x+1, y)].BIN != 999 && chipDictionary[(x, y+1)].BIN != 999&& chipDictionary[(x+1, y+1)].BIN != 999)
                        {
                            averagedChip.VF1 = chips.Average(c => c.VF1);
                        }

                        await Task.Run(() =>
                        {
                            lock (lockObject)
                            {
                                using (StreamWriter sw = new StreamWriter(outputCsvFile, true, Encoding.UTF8))
                                {

                                    sw.WriteLineAsync(string.Join(",", fieldOrder.Keys
                                        .Select(key =>
                                        {
                                            var value = typeof(Chip).GetProperty(key).GetValue(averagedChip);
                                            return Convert.ToDouble(value) != -100000 ? value.ToString() : "";
                                        })));
                                }
                            }
                        });

                    }
                }
            }

            try
            {
                // 读取CSV文件内容
                string[] liness = File.ReadAllLines(outputCsvFile);

                // 修改TotalTested列的数据
                for (int i = 0; i < lines.Length; i++)
                {
                    string[] columns = liness[i].Split(',');
                    if (columns.Length > 0 && columns[0] == "TotalTested")
                    {
                        // 假设TotalTested列是第二列（索引为1），修改为newValue
                        columns[2] = index.ToString();
                        liness[i] = string.Join(",", columns);
                        break; // 找到并修改第一次出现的TotalTested后退出循环
                    }
                }

                // 将修改后的内容写回CSV文件
                File.WriteAllLinesAsync(outputCsvFile, liness, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                MessageBox.Show("错误：" + ex.Message);
            }

            await Dispatcher.InvokeAsync(() =>
            {
                lock (parameterlockObject)
                {
                    parameterListBox.Items.Add(filename + " 计算完成!");
                    // 滚动到最新项
                    parameterListBox.ScrollIntoView(parameterListBox.Items[parameterListBox.Items.Count - 1]);
                }
            });
        }



        void initalBinList(List<BinData> binDataList)
        {
            foreach (BinData binDataTmp in binDataList)
            {
                binDataTmp.chipNum = 0;
            }
        }

        private async void LoadFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            double vf1fixNum = 1;
            double lop1fixNum = 1;
            openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            vf1fixNum = Convert.ToDouble(vf1TextBox.Text);
            lop1fixNum = Convert.ToDouble(lop1TextBox.Text);
            string outputFolder = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutputFolder");
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
            initalBinList(binDataList);
            totalChipNum = 0;
            if (openFileDialog.ShowDialog() == true)
            {
                DateTime startTime = DateTime.Now; // 记录开始时间

                List<Task> tasks = new List<Task>(); // 声明 tasks 列表


                // 尝试打开文件，如果文件已经被打开会引发 IOException 异常
                foreach (string filename in openFileDialog.FileNames)
                {
                    //string output_csv_file = System.IO.Path.Combine(outputFolder, filename);

                    string output_csv_file = System.IO.Path.Combine(outputFolder, System.IO.Path.GetFileNameWithoutExtension(filename) + ".csv");
                    tasks.Add(Task.Run(() => ProcessFile(filename, output_csv_file, vf1fixNum, lop1fixNum))); // 使用多线程处理文件
                    if (breakFlag)
                    {
                        break;
                    }
                }
                await Task.WhenAll(tasks); // 等待所有任务完成

                if (!breakFlag)
                {
                    DateTime endTime = DateTime.Now; // 记录结束时间
                    TimeSpan totalTime = endTime - startTime; // 计算运行时间
                    MessageBox.Show($"文件导入成功！总共耗时：{totalTime.TotalSeconds} 秒");
                }

            }
            else
            {
                MessageBox.Show("请输入文件！");
            }
        }
    }
}
