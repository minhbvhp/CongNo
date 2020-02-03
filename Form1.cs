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
using System.Globalization;


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
            //Đọc file Sunweb
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

                    //Đổ dữ liệu vào Database
                    int row;
                    String loai_ct;
                    String no_co;
                    String so_bk;
                    String mst;
                    String cong_ty1;
                    String cong_ty2;
                    String ky_hieu_hd;
                    String so_hoa_don;
                    String ngay_hoa_don;
                    String ma_nv;
                    String ma_phong;
                    String so_tien;
                    String han_tt;
                    String ngay_ct;
                    String user;

                    String db_name = "2019.accdb";
                    String db_path = Environment.CurrentDirectory + "\\";
                    String connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" + db_path + db_name;
                    OleDbConnection conn = new OleDbConnection(connectionString);
                    OleDbCommand dbCmd = new OleDbCommand();
                    
                    try
                    {
                        conn.Open();
                    }
                    catch (OleDbException exp)
                    {
                        MessageBox.Show("Database Error: " + exp.Message.ToString());
                    }

                    OleDbCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "INSERT INTO draft " +
                            "VALUES (@dong, @loai_ct, @no_co, @so_bk, @mst, @cong_ty1, " +
                            "@cong_ty2, @ky_hieu_hd, @so_hoa_don, @ngay_hoa_don, @ma_nv, " +
                            "@ma_phong, @so_tien, @han_tt, @ngay_ct, user)";

                    for (row = 2; row < lastRow; row++)
                    {
                        loai_ct = NullToString(worksheet.Cells["A" + row].Value);
                        no_co = NullToString(worksheet.Cells["D" + row].Value);
                        so_bk = NullToString(worksheet.Cells["E" + row].Value);
                        mst = NullToString(worksheet.Cells["K" + row].Value);
                        cong_ty1 = NullToString(worksheet.Cells["F" + row].Value);
                        cong_ty2 = NullToString(worksheet.Cells["G" + row].Value);
                        ky_hieu_hd = NullToString(worksheet.Cells["L" + row].Value);
                        so_hoa_don = NullToString(worksheet.Cells["M" + row].Value);
                        ngay_hoa_don = NullToString(worksheet.Cells["N" + row].Value);
                        ma_nv = NullToString(worksheet.Cells["I" + row].Value);
                        ma_phong = NullToString(worksheet.Cells["J" + row].Value);
                        so_tien = NullToString(worksheet.Cells["C" + row].Value);
                        han_tt = NullToString(worksheet.Cells["H" + row].Value);
                        ngay_ct = NullToString(worksheet.Cells["B" + row].Value);
                        user = NullToString(worksheet.Cells["O" + row].Value);


                        cmd.Parameters.Add("@dong", row.ToString());
                        cmd.Parameters.Add("@loai_ct", loai_ct);
                        cmd.Parameters.Add("@no_co", no_co);
                        cmd.Parameters.Add("@so_bk", so_bk);
                        cmd.Parameters.Add("@mst", mst);
                        cmd.Parameters.Add("@cong_ty1", cong_ty1);
                        cmd.Parameters.Add("@cong_ty2", cong_ty2);
                        cmd.Parameters.Add("@ky_hieu_hd", ky_hieu_hd);
                        cmd.Parameters.Add("@so_hoa_don", so_hoa_don);
                        cmd.Parameters.Add("@ngay_hoa_don", ngay_hoa_don);
                        cmd.Parameters.Add("@ma_nv", ma_nv);
                        cmd.Parameters.Add("@ma_phong", ma_phong);
                        cmd.Parameters.Add("@so_tien", so_tien);
                        cmd.Parameters.Add("@han_tt", han_tt);
                        cmd.Parameters.Add("@ngay_ct", ngay_ct);
                        cmd.Parameters.Add("@user", user);

                        cmd.ExecuteNonQuery();
                    }
                    

                    /*
                    for (row = 2; row < lastRow; row++)
                    {
                        
                        loai_ct = NullToString(worksheet.Cells["A" + row].Value);
                        no_co = NullToString(worksheet.Cells["D" + row].Value);
                        so_bk = NullToString(worksheet.Cells["E" + row].Value);
                        mst = NullToString(worksheet.Cells["K" + row].Value);
                        cong_ty1 = NullToString(worksheet.Cells["F" + row].Value);
                        cong_ty2 = NullToString(worksheet.Cells["G" + row].Value);
                        ky_hieu_hd = NullToString(worksheet.Cells["L" + row].Value);
                        so_hoa_don = NullToString(worksheet.Cells["M" + row].Value);
                        ngay_hoa_don = NullToDateTime(worksheet.Cells["N" + row].Value);
                        ma_nv = NullToString(worksheet.Cells["I" + row].Value);
                        ma_phong = NullToString(worksheet.Cells["J" + row].Value);
                        so_tien = NullToNumber(worksheet.Cells["C" + row].Value);
                        han_tt = NullToDateTime(worksheet.Cells["H" + row].Value);
                        ngay_ct = NullToDateTime(worksheet.Cells["B" + row].Value);
                        user = NullToString(worksheet.Cells["O" + row].Value);
                        
                        _2019DataSetTableAdapters.draftTableAdapter draftTableAdapter = new _2019DataSetTableAdapters.draftTableAdapter();
                        draftTableAdapter.Insert(row, loai_ct, no_co, so_bk, mst, cong_ty1, cong_ty2,
                            ky_hieu_hd, so_hoa_don, ngay_hoa_don, ma_nv, ma_phong, so_tien, han_tt, ngay_ct, user);

                        uploadProgress.Value = row / lastRow * 100;
                    }
                    */

                    if (conn.State == ConnectionState.Open)
                        conn.Close();
                    MessageBox.Show("Đã upload dữ liệu xong " + (lastRow - 1).ToString() + " bản ghi.");
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
            DateTime test = DateTime.ParseExact("01/01/2019", "dd/MM/yyyy", CultureInfo.InvariantCulture);
            MessageBox.Show(test.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture));
        }

        static string NullToString(object Value)
        {
            return Value == null ? "" : Value.ToString();
        }

        static double NullToNumber(object Value)
        {
            return Value == null ? 0 : (double)Value;
        }
        static DateTime NullToDateTime(object Value)
        {
            return Value == null ? Convert.ToDateTime("01/01/1900") : Convert.ToDateTime(Value);
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

                        worksheet.Cells["A:O"].Style.Font.Size = 8;
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
