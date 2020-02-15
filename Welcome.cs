using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace CongNo
{
    public partial class Welcome : Form
    {
        public void PickDB(string[] files)
        {
            comboNam.Items.Clear();

            if (files.Length > 0)
            {
                comboNam.Visible = true;
                numericNam.Visible = false;

                String name;
                foreach (string file in files)
                {
                    name = Path.GetFileNameWithoutExtension(file);
                    comboNam.Items.Add(name);
                }

                start.Text = "Báo cáo";
            }
            else
            {
                comboNam.Visible = false;
                numericNam.Visible = true;
                start.Text = "Tạo dữ liệu năm mới";

                MessageBox.Show("Chưa có dữ liệu, hãy tạo dữ liệu năm mới", "" , MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        public Welcome()
        {
            InitializeComponent();
        }

        private void Start_Click(object sender, EventArgs e)
        {
            Form form1 = new Form1();
            this.Hide();
            form1.Show();

            Program.OpenDetailFormOnClose = true;
            this.Close();
        }

        private void Welcome_Load(object sender, EventArgs e)
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(Environment.CurrentDirectory + @"\Database");

            if (!directoryInfo.Exists)
                Directory.CreateDirectory(Environment.CurrentDirectory + @"\Database");

            String[] files = Directory.GetFiles(Environment.CurrentDirectory + @"\Database\", "*.mdb");

            PickDB(files);
        }
    }
}
