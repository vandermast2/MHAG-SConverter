using System;
using System.IO;
using System.Windows.Forms;

namespace MHAGConverter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var filePath = string.Empty;

            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "e:\\";
                openFileDialog.Filter = @"xlsx files (*.xlsx)|*.xlsx";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (var reader = new StreamReader(fileStream))
                    {
                        reader.ReadToEnd();
                    }
                }
            }
            textBox1.Text = filePath;

            //;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show($@"Is completed successful: {new Converters().Ver1(textBox1.Text)}",
                @"STATUS",
                buttons: MessageBoxButtons.OK);
        }
    }

}
