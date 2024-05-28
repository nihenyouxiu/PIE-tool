using System.Collections.ObjectModel;
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
    /// Interaction logic for MainWindow.xaml
    /// </summary>
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
    }
}