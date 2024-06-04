using System;
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
    /// Page2.xaml 的交互逻辑
    /// </summary>
    /// 
    public class BinData
    {
        public double binIdx { get; set; }
        public double chipNum { get; set; }
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

        public readonly object Lock = new object();
        public override string ToString()
        {
            return $"Bin: {binIdx}, VF1Min: {VF1Min}, " +
                   $"VF1Max: {VF1Max}, VF2Min: {VF2Min}, VF2Max: {VF2Max}, " +
                   $"VF3Min: {VF3Min}, VF3Max: {VF3Max}, VF4Min: {VF4Min}, " +
                   $"VF4Max: {VF4Max}, VZ1Min: {VZ1Min}, VZ1Max: {VZ1Max}, " +
                   $"IRMin: {IRMin}, IRMax: {IRMax}, HW1Min: {HW1Min}, " +
                   $"HW1Max: {HW1Max}, LOP1Min: {LOP1Min}, LOP1Max: {LOP1Max}, " +
                   $"WLP1Min: {WLP1Min}, WLP1Max: {WLP1Max}, WLD1Min: {WLD1Min}, " +
                   $"WLD1Max: {WLD1Max}, IR1Min: {IR1Min}, IR1Max: {IR1Max}, " +
                   $"VFDMin: {VFDMin}, VFDMax: {VFDMax}, DVFMin: {DVFMin}, " +
                   $"DVFMax: {DVFMax}, IR2Min: {IR2Min}, IR2Max: {IR2Max}, " +
                   $"WLC1Min: {WLC1Min}, WLC1Max: {WLC1Max}, VF5Min: {VF5Min}, " +
                   $"VF5Max: {VF5Max}, VF6Min: {VF6Min}, VF6Max: {VF6Max}, " +
                   $"VF7Min: {VF7Min}, VF7Max: {VF7Max}, VF8Min: {VF8Min}, " +
                   $"VF8Max: {VF8Max}, DVF1Min: {DVF1Min}, DVF1Max: {DVF1Max}, " +
                   $"DVF2Min: {DVF2Min}, DVF2Max: {DVF2Max}, VZ2Min: {VZ2Min}, " +
                   $"VZ2Max: {VZ2Max}, VZ3Min: {VZ3Min}, VZ3Max: {VZ3Max}, " +
                   $"VZ4Min: {VZ4Min}, VZ4Max: {VZ4Max}, VZ5Min: {VZ5Min}, " +
                   $"VZ5Max: {VZ5Max}, IR3Min: {IR3Min}, IR3Max: {IR3Max}, " +
                   $"IR4Min: {IR4Min}, IR4Max: {IR4Max}, IR5Min: {IR5Min}, " +
                   $"IR5Max: {IR5Max}, IR6Min: {IR6Min}, IR6Max: {IR6Max}, " +
                   $"IFMin: {IFMin}, IFMax: {IFMax}, IF1Min: {IF1Min}, " +
                   $"IF1Max: {IF1Max}, IF2Min: {IF2Min}, IF2Max: {IF2Max}, " +
                   $"LOP2Min: {LOP2Min}, LOP2Max: {LOP2Max}, WLP2Min: {WLP2Min}, " +
                   $"WLP2Max: {WLP2Max}, WLD2Min: {WLD2Min}, WLD2Max: {WLD2Max}, " +
                   $"HW2Min: {HW2Min}, HW2Max: {HW2Max}, WLC2Min: {WLC2Min}, " +
                   $"WLC2Max: {WLC2Max}";
        }
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
    }
    public partial class Page2 : Page
    {
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

        public Page2()
        {
            InitializeComponent();
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

            try
            {
                using (StreamReader reader = new StreamReader(filename))
                {

                    while (!reader.EndOfStream)
                    {
                        string[] values = reader.ReadLine().Split(',');
                        string firstValue = values[0];
                        bool isFirstValueAllDigits = Regex.IsMatch(firstValue, @"^\d+$");
                        if (isFirstValueAllDigits && values.Length >= 56)
                        {
                            Chip chipData = new Chip();
                            chipData.TEST = !string.IsNullOrEmpty(values[0]) ? Convert.ToDouble(values[0]) : -100000;
                            chipData.BIN = !string.IsNullOrEmpty(values[1]) ? 999 : -100000;
                            chipData.VF1 = !string.IsNullOrEmpty(values[2]) ? Convert.ToDouble(values[2]) * vf1fixNum : -100000;
                            chipData.VF2 = !string.IsNullOrEmpty(values[3]) ? Convert.ToDouble(values[3]) : -100000;
                            chipData.VF3 = !string.IsNullOrEmpty(values[4]) ? Convert.ToDouble(values[4]) : -100000;
                            chipData.VF4 = !string.IsNullOrEmpty(values[5]) ? Convert.ToDouble(values[5]) : -100000;
                            chipData.VF5 = !string.IsNullOrEmpty(values[6]) ? Convert.ToDouble(values[6]) : -100000;
                            chipData.VF6 = !string.IsNullOrEmpty(values[7]) ? Convert.ToDouble(values[7]) : -100000;
                            chipData.VF = !string.IsNullOrEmpty(values[9]) ? Convert.ToDouble(values[9]) : -100000;
                            chipData.VZ1 = !string.IsNullOrEmpty(values[11]) ? Convert.ToDouble(values[11]) : -100000;
                            chipData.VZ2 = !string.IsNullOrEmpty(values[12]) ? Convert.ToDouble(values[12]) : -100000;
                            chipData.IR = !string.IsNullOrEmpty(values[13]) ? Convert.ToDouble(values[13]) : -100000;
                            chipData.LOP1 = !string.IsNullOrEmpty(values[14]) ? Convert.ToDouble(values[14]) * lop1fixNum : -100000;
                            chipData.LOP2 = !string.IsNullOrEmpty(values[15]) ? Convert.ToDouble(values[15]) : -100000;
                            chipData.LOP3 = !string.IsNullOrEmpty(values[16]) ? Convert.ToDouble(values[16]) : -100000;
                            chipData.WLP1 = !string.IsNullOrEmpty(values[17]) ? Convert.ToDouble(values[17]) : -100000;
                            chipData.WLD1 = !string.IsNullOrEmpty(values[18]) ? Convert.ToDouble(values[18]) : -100000;
                            chipData.WLC1 = !string.IsNullOrEmpty(values[19]) ? Convert.ToDouble(values[19]) : -100000;
                            chipData.HW1 = !string.IsNullOrEmpty(values[20]) ? Convert.ToDouble(values[20]) : -100000;
                            chipData.WLP2 = !string.IsNullOrEmpty(values[27]) ? Convert.ToDouble(values[27]) : -100000;
                            chipData.WLD2 = !string.IsNullOrEmpty(values[28]) ? Convert.ToDouble(values[28]) : -100000;
                            chipData.WLC2 = !string.IsNullOrEmpty(values[29]) ? Convert.ToDouble(values[29]) : -100000;
                            chipData.HW2 = !string.IsNullOrEmpty(values[30]) ? Convert.ToDouble(values[30]) : -100000;
                            chipData.VF7 = !string.IsNullOrEmpty(values[36]) ? Convert.ToDouble(values[36]) : -100000;
                            chipData.VF8 = !string.IsNullOrEmpty(values[37]) ? Convert.ToDouble(values[37]) : -100000;
                            chipData.IR3 = !string.IsNullOrEmpty(values[38]) ? Convert.ToDouble(values[38]) : -100000;
                            chipData.IR4 = !string.IsNullOrEmpty(values[39]) ? Convert.ToDouble(values[39]) : -100000;
                            chipData.IR5 = !string.IsNullOrEmpty(values[40]) ? Convert.ToDouble(values[40]) : -100000;
                            chipData.IR6 = !string.IsNullOrEmpty(values[41]) ? Convert.ToDouble(values[41]) : -100000;
                            chipData.VZ3 = !string.IsNullOrEmpty(values[42]) ? Convert.ToDouble(values[42]) : -100000;
                            chipData.VZ4 = !string.IsNullOrEmpty(values[43]) ? Convert.ToDouble(values[43]) : -100000;
                            chipData.VZ5 = !string.IsNullOrEmpty(values[44]) ? Convert.ToDouble(values[44]) : -100000;
                            chipData.IF = !string.IsNullOrEmpty(values[45]) ? Convert.ToDouble(values[45]) : -100000;
                            chipData.IF1 = !string.IsNullOrEmpty(values[46]) ? Convert.ToDouble(values[46]) : -100000;
                            chipData.IF2 = !string.IsNullOrEmpty(values[47]) ? Convert.ToDouble(values[47]) : -100000;
                            chipData.IR1 = !string.IsNullOrEmpty(values[50]) ? Convert.ToDouble(values[50]) : -100000;
                            chipData.IR2 = !string.IsNullOrEmpty(values[51]) ? Convert.ToDouble(values[51]) : -100000;
                            chipData.VFD = !string.IsNullOrEmpty(values[10]) ? Convert.ToDouble(values[10]) : -100000;
                            chipData.DVF = (dvfMax == dvfMin ? -100000 : chipData.VF2 - chipData.VF3);
                            chipData.DVF1 = (dvf1Max == dvf1Min ? -100000 : chipData.VF6 - chipData.VF4);
                            chipData.DVF2 = (dvf1Max == dvf1Min ? -100000 : chipData.VF8 - chipData.VF6);
                            chipList.Add(chipData);
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

            BinData binDataFailTmp = binDataList.FirstOrDefault(item => item.binIdx == 999);
            int waferidchipnum = 0;
            if (chipList.Any())
            {
                foreach (Chip chip in chipList)
                {
                    bool flag = false;

                    foreach (BinData binDataTmp in binDataList)
                    {
                        if (ValidateAgainstBinData(chip, binDataTmp))
                        {
                            lock (binDataTmp.Lock)
                            {
                                binDataTmp.chipNum++;
                                chip.BIN = binDataTmp.binIdx;
                            }
                            flag = true;
                            break;
                        }
                    }

                    if (!flag)
                    {
                        lock (binDataFailTmp?.Lock)
                        {
                            if (binDataFailTmp != null)
                            {
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

            await Task.Run(() =>
            {
                lock (lockObject)
                {
                    using (StreamWriter sw = new StreamWriter(outputCsvFile, true, Encoding.UTF8))
                    {
                        sw.WriteLineAsync(System.IO.Path.GetFileNameWithoutExtension(filename) + "," + waferidchipnum.ToString());
                    }
                }
            });

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
            string output_csv_file = System.IO.Path.Combine(outputFolder, "out.csv");
            using (StreamWriter sw = new StreamWriter(output_csv_file, true, Encoding.UTF8))
            {
                sw.WriteLine("waferid,chipNum");
            }
            initalBinList(binDataList);
            totalChipNum = 0;
            if (openFileDialog.ShowDialog() == true)
            {
                DateTime startTime = DateTime.Now; // 记录开始时间

                List<Task> tasks = new List<Task>(); // 声明 tasks 列表


                // 尝试打开文件，如果文件已经被打开会引发 IOException 异常
                foreach (string filename in openFileDialog.FileNames)
                {
                    //string output_csv_file = System.IO.Path.Combine(outputFolder, System.IO.Path.GetFileName(filename));
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
