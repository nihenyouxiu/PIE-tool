﻿using System.Collections.ObjectModel;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace 产出分布计算
{
    /// <summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void FramePage1_Loaded(object sender, RoutedEventArgs e)
        {
            FramePage1.Navigate(new Page1());
        }

        private void FramePage2_Loaded(object sender, RoutedEventArgs e)
        {
            FramePage2.Navigate(new Page2());
        }
        private void FramePage3_Loaded(object sender, RoutedEventArgs e)
        {
            FramePage3.Navigate(new Page3());
        }

        private void FramePage4_Loaded(object sender, RoutedEventArgs e)
        {
            FramePage4.Navigate(new Page4());
        }

        private void FramePage5_Loaded(object sender, RoutedEventArgs e)
        {
            FramePage5.Navigate(new Page5());
        }

    }

}