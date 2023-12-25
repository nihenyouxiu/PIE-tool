using System;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace SimpleApp
{
    public partial class MainForm : Form
    {
        private TextBox[] textInputs;
        private Button concatButton;
        private TextBox displayTextBox;

        public MainForm()
        {
            InitializeComponent();
            InitializeUI();
        }

        private void InitializeUI()
        {
            // 初始化四个文本输入框
            textInputs = new TextBox[4];
            for (int i = 0; i < 4; i++)
            {
                textInputs[i] = new TextBox();
                textInputs[i].Location = new System.Drawing.Point(20, 20 + i * 30);
                textInputs[i].Size = new System.Drawing.Size(150, 25);
                this.Controls.Add(textInputs[i]);
            }

            // 初始化按钮
            concatButton = new Button();
            concatButton.Text = "Concatenate";
            concatButton.Location = new System.Drawing.Point(20, 140);
            concatButton.Size = new System.Drawing.Size(150, 30);
            concatButton.Click += ConcatButton_Click;
            this.Controls.Add(concatButton);

            // 初始化显示框
            displayTextBox = new TextBox();
            displayTextBox.Multiline = true;
            displayTextBox.ReadOnly = true;
            displayTextBox.Location = new System.Drawing.Point(200, 20);
            displayTextBox.Size = new System.Drawing.Size(200, 150);
            this.Controls.Add(displayTextBox);
        }

        private void ConcatButton_Click(object sender, EventArgs e)
        {
            // 当按钮被点击时，将四个文本框中的文本拼接起来并显示在显示框中
            string concatenatedText = string.Join(" ", textInputs.Select(tb => tb.Text));
            displayTextBox.Text = concatenatedText;
        }

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
