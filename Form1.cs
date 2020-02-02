using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;
using System.Data.OleDb;
using OfficeOpenXml.Style;


namespace CongNo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void upload_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "File mẫu|Mau lay du lieu Sunweb.xlsx";
            openFileDialog1.FileName = "Mau lay du lieu Sunweb";
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                String inputPath = Path.GetFullPath(openFileDialog1.FileName);
                var importExcel = new FileInfo(inputPath);
                using (var package = new ExcelPackage(importExcel))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    int lastRow = (int)worksheet.Dimension.End.Row;

                    StreamWriter sw = new StreamWriter(Environment.CurrentDirectory + @"\Database\Input.txt");
                    for (int i = 2; i <= lastRow; i++)
                    {
                        var loai_ct = worksheet.Cells["A" + i].Value;
                        var no_co = worksheet.Cells["D" + i].Value;
                        var so_bk = worksheet.Cells["E" + i].Value;
                        var mst = worksheet.Cells["J" + i].Value;
                        var cong_ty1 = worksheet.Cells["E" + i].Value;
                        var cong_ty2 = worksheet.Cells["F" + i].Value;
                        var ky_hieu_hd = worksheet.Cells["K" + i].Value;
                        var so_hoa_don = worksheet.Cells["L" + i].Value;
                        var ngay_hoa_don = worksheet.Cells["M" + i].Value;
                        var ma_nv = worksheet.Cells["H" + i].Value;
                        var ma_phong = worksheet.Cells["I" + i].Value;
                        var so_tien = worksheet.Cells["C" + i].Value;
                        var han_tt = worksheet.Cells["G" + i].Value;
                        var ngay_ct = worksheet.Cells["B" + i].Value;
                        var user = worksheet.Cells["N" + i].Value;

                        sw.Write("{0} - {1} - {2} - {3} - {4} - {5} - {6} - {7} - {8} - {9}" +
                            " - {10} - {11} - {12} - {13} - {14}" + Environment.NewLine, NullToString(loai_ct), NullToString(no_co), NullToString(so_bk), NullToString(mst), NullToString(cong_ty1),
                            NullToString(cong_ty2), NullToString(ky_hieu_hd), NullToString(so_hoa_don),
                            NullToString(ngay_hoa_don), NullToString(ma_nv), NullToString(ma_phong), NullToString(so_tien),
                            NullToString(han_tt), NullToString(ngay_ct), NullToString(user));

                        uploadProgress.Value = i / lastRow * 100;
                    }
                    sw.Close();
                    MessageBox.Show("Đã upload dữ liệu xong");
                }
            }
                
                
            

        }

        private void Report_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Excel phiên bản 2007 trở lên|*.xlsx";
                if (saveDialog.ShowDialog() != DialogResult.Cancel)
                {
                    string exportFilePath = saveDialog.FileName;
                    var newFile = new FileInfo(exportFilePath);
                    using (var package = new ExcelPackage(newFile))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("NewSheet1");
                        worksheet.Cells["A1"].Value = "Xin chào mẹ Thương ngố";
                        package.Save();
                    }
                }
            }
        }

        private void Search_Click(object sender, EventArgs e)
        {
            OpenDb("2019");
        }

        static string NullToString(object Value)
        {
            return Value == null ? "" : Value.ToString();
        }
        public void OpenDb(string dbname)
        {
            String db_name = dbname + ".accdb";
            String db_path = Environment.CurrentDirectory + @"\Database\";
            String connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" + db_path + db_name;

            //Mở Database, tạo mới nếu chưa có
            if (!File.Exists(db_path + db_name))
            {
                ADOX.Catalog cat = new ADOX.Catalog();
                cat.Create(connectionString);

                ADODB.Connection con = cat.ActiveConnection as ADODB.Connection;

                //Tạo bảng
                String createCustomers = @"CREATE TABLE customers(mst VARCHAR(200) PRIMARY KEY NOT NULL, " +
                    "cong_ty VARCHAR(200))";
                OleDbConnection conn = new OleDbConnection(connectionString);
                OleDbCommand dbCmd = new OleDbCommand();

                try
                {
                    conn.Open();
                    MessageBox.Show(createCustomers);
                    dbCmd.Connection = conn;
                    dbCmd.CommandText = createCustomers;
                    dbCmd.ExecuteNonQuery();
                    MessageBox.Show("Table created");

                    String qery = @"CREATE PROCEDURE test as SELECT * FROM customers";
                    OleDbCommand createQuery = new OleDbCommand(qery, conn);
                    createQuery.ExecuteNonQuery();
                    MessageBox.Show("Query created");
                }
                catch (OleDbException exp)
                {
                    MessageBox.Show("Database Error: " + exp.Message.ToString());
                }
                finally
                {
                    if (conn.State == ConnectionState.Open)
                        conn.Close();
                }
                if (con != null)
                    con.Close();
            }
            else
            {

            }
        }

        //Export file "Mau lay du lieu Sunweb.xlsx"
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            String selectedFolder = "";
            FolderBrowserDialog sunwebLocation = new FolderBrowserDialog();
            sunwebLocation.SelectedPath = selectedFolder;
            DialogResult result = sunwebLocation.ShowDialog();
            if (result == DialogResult.OK)
            {
                String fileName = "Mau lay du lieu Sunweb.xlsx";
                selectedFolder = sunwebLocation.SelectedPath;

                var newFile = new FileInfo(selectedFolder + "\\" + fileName);

                //Tạo mẫu
                if (!File.Exists(newFile.ToString()))
                {
                    using (var package = new ExcelPackage(newFile))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Du lieu Sunweb");
                        worksheet.Cells["A1:O1"].Style.Font.Name = "Calibri";
                        worksheet.Cells["A1:O1"].Style.Font.Bold = true;

                        worksheet.Column(1).Width = 6;
                        worksheet.Column(2).Width = 9;
                        worksheet.Column(3).Width = 12;
                        worksheet.Column(4).Width = 5;
                        worksheet.Column(5).Width = 12;
                        worksheet.Column(6).Width = 35;
                        worksheet.Column(7).Width = 30;
                        worksheet.Column(8).Width = 11;
                        worksheet.Column(9).Width = 8;
                        worksheet.Column(10).Width = 8.5;
                        worksheet.Column(11).Width = 13;
                        worksheet.Column(12).Width = 13;
                        worksheet.Column(13).Width = 9;
                        worksheet.Column(14).Width = 15;
                        worksheet.Column(15).Width = 15.20;

                        worksheet.Cells["A:O"].Style.Font.Size = 9;
                        worksheet.Cells["A:O"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["A:O"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["A:O"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["A:O"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        worksheet.Cells["A1"].Value = "LOAICT";
                        worksheet.Cells["B1"].Value = "NGAYCT";
                        worksheet.Cells["C1"].Value = "SOTIENQUYDOI";
                        worksheet.Cells["D1"].Value = "NOCO";
                        worksheet.Cells["E1"].Value = "SOTHAMCHIEU";
                        worksheet.Cells["F1"].Value = "GNRL_DESCR_01";
                        worksheet.Cells["G1"].Value = "GNRL_DESCR_02";
                        worksheet.Cells["H1"].Value = "NGAYDAOHAN";
                        worksheet.Cells["I1"].Value = "T2";
                        worksheet.Cells["J1"].Value = "T3";
                        worksheet.Cells["K1"].Value = "MASOTHUE";
                        worksheet.Cells["L1"].Value = "KYHIEUHOADON";
                        worksheet.Cells["M1"].Value = "SOHOADON";
                        worksheet.Cells["N1"].Value = "NGAYHOADONGOC";
                        worksheet.Cells["O1"].Value = "USERNHAP";

                        package.Save();
                    }
                    MessageBox.Show("Đã tải xuống File mẫu", "Thành công!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Bạn đã tải xuống File này rồi", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
