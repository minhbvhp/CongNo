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
                comboNam.SelectedIndex = 0;

                start.Text = "Báo cáo";
            }
            else
            {
                comboNam.Visible = false;
                numericNam.Visible = true;
                start.Text = "Tạo dữ liệu năm mới";
            }
        }
        public Welcome()
        {
            InitializeComponent();
        }

        private void Start_Click(object sender, EventArgs e)
        {
            if (comboNam.Visible)
                Program.DbYear = comboNam.SelectedItem.ToString();
            else
                Program.DbYear = numericNam.Value.ToString();

            String DBFileName = Program.DbYear + ".mdb";
            FileInfo fileInfo = new FileInfo(Environment.CurrentDirectory + @"\Database\" + DBFileName);

            if (!fileInfo.Exists)
            {
                Directory.CreateDirectory(Environment.CurrentDirectory + @"\Database");
                using (Stream s = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("CongNo.DB.mdb"))
                {
                    using (FileStream ResourceFile = new FileStream(fileInfo.ToString(), FileMode.Create, FileAccess.Write))
                    {
                        s.CopyTo(ResourceFile);
                    }
                }

                String db_file = fileInfo.ToString();

                DAO.DBEngine dBEngine = new DAO.DBEngine();
                DAO.Database db;
                db = dBEngine.OpenDatabase(db_file);
                String queryName = "cong_no_draft";
                String querySql = String.Format("SELECT invoice.ngay_ct, invoice.ngay_hoa_don, department.ma_phong," +
                    " department.ten_phong, customers.mst, customers.cong_ty, " +
                    "invoice.ki_hieu_hoa_don, invoice.so_hoa_don, invoice.han_thanh_toan, revenue.ma_nv, revenue.user_nhap, invoice.kenh_kt, " +
                    "IIf(Year(invoice.ngay_ct)<{0},invoice.so_tien_phat_sinh) AS du_dau_ky, IIf(Year(invoice.ngay_ct)={0} " +
                    "And Month(invoice.ngay_ct)=1,invoice.so_tien_phat_sinh) AS no1, IIf(Year(invoice.ngay_ct)={0} And " +
                    "Month(invoice.ngay_ct)=2,invoice.so_tien_phat_sinh) AS no2, IIf(Year(invoice.ngay_ct)={0} And " +
                    "Month(invoice.ngay_ct)=3,invoice.so_tien_phat_sinh) AS no3, IIf(Year(invoice.ngay_ct)={0} And " +
                    "Month(invoice.ngay_ct)=4,invoice.so_tien_phat_sinh) AS no4, IIf(Year(invoice.ngay_ct)={0} And " +
                    "Month(invoice.ngay_ct)=5,invoice.so_tien_phat_sinh) AS no5, IIf(Year(invoice.ngay_ct)={0} And " +
                    "Month(invoice.ngay_ct)=6,invoice.so_tien_phat_sinh) AS no6, IIf(Year(invoice.ngay_ct)={0} And " +
                    "Month(invoice.ngay_ct)=7,invoice.so_tien_phat_sinh) AS no7, IIf(Year(invoice.ngay_ct)={0} And " +
                    "Month(invoice.ngay_ct)=8,invoice.so_tien_phat_sinh) AS no8, IIf(Year(invoice.ngay_ct)={0} And " +
                    "Month(invoice.ngay_ct)=9,invoice.so_tien_phat_sinh) AS no9, IIf(Year(invoice.ngay_ct)={0} And " +
                    "Month(invoice.ngay_ct)=10,invoice.so_tien_phat_sinh) AS no10, IIf(Year(invoice.ngay_ct)={0} And " +
                    "Month(invoice.ngay_ct)=11,invoice.so_tien_phat_sinh) AS no11, IIf(Year(invoice.ngay_ct)={0} And " +
                    "Month(invoice.ngay_ct)=12,invoice.so_tien_phat_sinh) AS no12 FROM department " +
                    "INNER JOIN((revenue INNER JOIN invoice ON (revenue.ki_hieu_hoa_don = invoice.ki_hieu_hoa_don)" +
                    " AND(revenue.so_hoa_don = invoice.so_hoa_don)) INNER JOIN customers ON invoice.mst = customers.mst)" +
                    " ON department.ma_phong = revenue.ma_phong ORDER BY invoice.ki_hieu_hoa_don, invoice.so_hoa_don;", Program.DbYear);

                DAO.QueryDef cong_no_draft = new DAO.QueryDef();
                cong_no_draft.Name = queryName;
                cong_no_draft.SQL = querySql;

                db.QueryDefs.Append(cong_no_draft);
                db.Close();
            }

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
