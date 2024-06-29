using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Windows;

namespace 产出分布计算
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    /// 
    public class MainViewModel : INotifyPropertyChanged
    {
        private bool _p1Checked;
        private bool _p2Checked;
        private bool _p3Checked;
        private string _textBox1Text;
        private string _textBox12Text;
        private ObservableCollection<string> _fileList;

        public ObservableCollection<string> FileList
        {
            get => _fileList;
            set
            {
                _fileList = value;
                OnPropertyChanged(nameof(FileList));
            }
        }

        public bool P1Checked
        {
            get => _p1Checked;
            set
            {
                _p1Checked = value;
                OnPropertyChanged(nameof(P1Checked));
            }
        }
        public bool P2Checked
        {
            get => _p2Checked;
            set
            {
                _p2Checked = value;
                OnPropertyChanged(nameof(P2Checked));
            }
        }
        public bool P3Checked
        {
            get => _p3Checked;
            set
            {
                _p3Checked = value;
                OnPropertyChanged(nameof(P3Checked));
            }
        }

        public string TextBox1Text
        {
            get => _textBox1Text;
            set
            {
                _textBox1Text = value;
                OnPropertyChanged(nameof(TextBox1Text));
            }
        }

        public string TextBox12Text
        {
            get => _textBox12Text;
            set
            {
                _textBox12Text = value;
                OnPropertyChanged(nameof(TextBox12Text));
            }
        }

        // Repeat the above properties for other controls

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
    public partial class App : Application
    {
        public static MainViewModel MainViewModel { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            MainViewModel = new MainViewModel();
        }
    }

}
