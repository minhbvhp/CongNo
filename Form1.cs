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
using OfficeOpenXml.Style;


namespace CongNo
{
    public partial class Form1 : Form
    {
        Dictionary<string, string[]> searchBy = new Dictionary<string, string[]>();
        static string NullToString(object Value)
        {
            return Value == null ? null : Value.ToString().Trim();
        }
        public String GhiChu(String soHoaDon, String khachHang, String maHoaDon)
        {
            String result = "";
            if (string.IsNullOrEmpty(soHoaDon) || string.IsNullOrWhiteSpace(soHoaDon))
                result =  "Thiếu số hóa đơn";
            if (string.IsNullOrEmpty(khachHang) || string.IsNullOrWhiteSpace(khachHang))
                result = "Thiếu MST và tên khách hàng";
            if (string.IsNullOrEmpty(maHoaDon) || string.IsNullOrWhiteSpace(maHoaDon))
                result = "Thiếu mã hóa đơn";

            return result;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            searchBy["Khách hàng"] = new string[] { "Mã số thuế", "Tên đơn vị" };
            searchBy["Phát sinh"] = new string[] { "Số hóa đơn", "Số tiền" };
            searchBy["Thu nợ"] = new string[] { "Số hóa đơn", "Số tiền" };

            foreach (String search in searchBy.Keys)
                categorySearch.Items.Add(search);

        }
        public Form1()
        {
            InitializeComponent();
        }

        private void upload_Click(object sender, EventArgs e)
        {
            notUploadList.Rows.Clear();
            searchList.Rows.Clear();

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
                    //Lấy tổng số bản ghi
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    int lastRow = (int)worksheet.Dimension.End.Row;

                    //Tạo connection
                    int row;

                    String db_name = "2019.mdb";
                    String db_path = Environment.CurrentDirectory + @"\Database\";
                    String db_file = db_path + db_name;

                    dao.DBEngine dBEngine = new dao.DBEngine();
                    dao.Database db;
                    dao.Recordset rs;

                    //Đổ dữ liệu vào Database
                    try
                    {
                        db = dBEngine.OpenDatabase(db_file);
                        db.Execute("draft_clear");
                        rs = db.OpenRecordset("draft");
                        dBEngine.BeginTrans();

                        for (row = 2; row < lastRow; row++)
                        {
                            rs.AddNew();
                            rs.Fields["dong"].Value = row;
                            rs.Fields["loai_ct"].Value = NullToString(worksheet.Cells["A" + row].Value);
                            rs.Fields["no_co"].Value = NullToString(worksheet.Cells["D" + row].Value);
                            rs.Fields["so_bk"].Value = NullToString(worksheet.Cells["E" + row].Value);
                            rs.Fields["mst"].Value = NullToString(worksheet.Cells["K" + row].Value);
                            rs.Fields["cong_ty1"].Value = NullToString(worksheet.Cells["F" + row].Value);
                            rs.Fields["cong_ty2"].Value = NullToString(worksheet.Cells["G" + row].Value);
                            rs.Fields["ky_hieu_hd"].Value = NullToString(worksheet.Cells["L" + row].Value);
                            rs.Fields["so_hoa_don"].Value = NullToString(worksheet.Cells["M" + row].Value);
                            rs.Fields["ngay_hoa_don"].Value = NullToString(worksheet.Cells["N" + row].Value);
                            rs.Fields["ma_nv"].Value = NullToString(worksheet.Cells["I" + row].Value);
                            rs.Fields["ma_phong"].Value = NullToString(worksheet.Cells["J" + row].Value);
                            rs.Fields["so_tien"].Value = NullToString(worksheet.Cells["C" + row].Value);
                            rs.Fields["han_tt"].Value = NullToString(worksheet.Cells["H" + row].Value);
                            rs.Fields["ngay_ct"].Value = NullToString(worksheet.Cells["B" + row].Value);
                            rs.Fields["user"].Value = NullToString(worksheet.Cells["O" + row].Value);
                            rs.Update();

                            uploadProgress.Value = row * 100 / lastRow;
                            uploadProgress.Refresh();
                            Application.DoEvents();
                        }

                        db.Execute("update_mst");
                        db.Execute("add_mst_to_customers");
                        db.Execute("invoice_filter");
                        db.Execute("update_invoice");
                        db.Execute("add_invoice");
                        db.Execute("update_paid");
                        db.Execute("add_paid");
                        db.Execute("update_revenue");
                        db.Execute("add_revenue");

                        uploadProgress.Value += 1;
                        uploadProgress.Refresh();

                        MessageBox.Show("Đã upload dữ liệu xong.", "Chúc mừng", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        //Liệt kê dữ liệu chưa được Upload
                        rs = db.OpenRecordset("not_upload");
                        if (!rs.EOF)
                            rs.MoveLast();

                        if (rs.RecordCount > 0)
                        {
                            String dong;
                            String ten_cong_ty;
                            String so_bk;
                            String ky_hieu_hd;
                            String so_hoa_don;
                            double so_tien;
                            String user;
                            String ghiChu;

                            if (!rs.BOF)
                                rs.MoveFirst();

                            while (!rs.EOF)
                            {
                                dong = Convert.ToString(rs.Fields["dong"].Value);
                                ten_cong_ty = Convert.ToString(rs.Fields["ten_cong_ty"].Value);
                                so_bk = Convert.ToString(rs.Fields["so_bk"].Value);
                                ky_hieu_hd = Convert.ToString(rs.Fields["ky_hieu_hd"].Value);
                                so_hoa_don = Convert.ToString(rs.Fields["so_hoa_don"].Value);
                                so_tien = rs.Fields["so_tien"].Value;
                                user = Convert.ToString(rs.Fields["user"].Value);
                                ghiChu = GhiChu(so_hoa_don, ten_cong_ty, ky_hieu_hd);

                                notUploadList.Rows.Add(dong, ten_cong_ty, so_bk, ky_hieu_hd, so_hoa_don, so_tien, user, ghiChu);
                                rs.MoveNext();
                            }
                        }

                        db.Execute("draft_clear");
                        db.Execute("invoice_draft_clear");
                        dBEngine.CommitTrans();
                        rs.Close();
                        db.Close();

                        String compactDbTemp = db_path + "temp.mdb";
                        String compactDbName = db_path + "2019.mdb";
                        dBEngine.CompactDatabase(db_file, compactDbTemp);
                        File.Delete(db_file);
                        File.Move(compactDbTemp, compactDbName);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Database error: " + ex.Message.ToString(), "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        private void CategorySearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            fieldSearch.Items.Clear();
            if (categorySearch.SelectedIndex > -1)
            {
                String field = searchBy.Keys.ElementAt(categorySearch.SelectedIndex);
                fieldSearch.Items.AddRange(searchBy[field]);

                switch (field)
                {
                    case "Khách hàng":
                        searchList.Columns["mst"].Visible = true;
                        searchList.Columns["ten_don_vi"].Visible = true;
                        searchList.Columns["ma_hoa_don"].Visible = false;
                        searchList.Columns["so_hoa_don"].Visible = false;
                        searchList.Columns["so_tien"].Visible = false;
                        searchList.Columns["recNumber"].Visible = true;
                        break;
                    case "Phát sinh":
                        searchList.Columns["mst"].Visible = true;
                        searchList.Columns["ten_don_vi"].Visible = false;
                        searchList.Columns["ma_hoa_don"].Visible = true;
                        searchList.Columns["so_hoa_don"].Visible = true;
                        searchList.Columns["so_tien"].Visible = true;
                        searchList.Columns["recNumber"].Visible = true;
                        break;
                    case "Thu nợ":
                        searchList.Columns["mst"].Visible = true;
                        searchList.Columns["ten_don_vi"].Visible = false;
                        searchList.Columns["ma_hoa_don"].Visible = true;
                        searchList.Columns["so_hoa_don"].Visible = true;
                        searchList.Columns["so_tien"].Visible = true;
                        searchList.Columns["recNumber"].Visible = true;
                        break;
                }
            }
        }

        private void Report_Click(object sender, EventArgs e)
        {
            String db_name = "2019.mdb";
            String db_path = Environment.CurrentDirectory + @"\Database\";
            String db_file = db_path + db_name;

            dao.DBEngine dBEngine = new dao.DBEngine();
            dao.Database db;
            dao.Recordset rs;

            db = dBEngine.OpenDatabase(db_file);
            rs = db.OpenRecordset("invoice");
            MessageBox.Show(rs.Name);

            MessageBox.Show(fieldSearch.Text);
        }

        private void Search_Click(object sender, EventArgs e)
        {
            searchList.Rows.Clear();
            String searchWhat = tbSearch.Text.Trim();

            String db_name = "2019.mdb";
            String db_path = Environment.CurrentDirectory + @"\Database\";
            String db_file = db_path + db_name;

            dao.DBEngine dBEngine = new dao.DBEngine();
            dao.Database db;
            dao.Recordset rs = null;

            /*
             * Toxic ways to avoid SQL injection
             * Don't think another ways.
             * Will optimize when find out.
             */

            try
            {
                db = dBEngine.OpenDatabase(db_file);
                String field = searchBy.Keys.ElementAt(categorySearch.SelectedIndex);

                String searchValue;
                Double searchDouble;
                bool matchSearchCondition = false;
                String mst = null;
                String ten_don_vi = null;
                String ma_hoa_don = null;
                String so_hoa_don = null;
                double so_tien = 0;
                double recordNumber = 0;
            
                switch (field)
                {
                    case "Khách hàng":
                        rs = db.OpenRecordset("customers");
                        break;
                    case "Phát sinh":
                        rs = db.OpenRecordset("invoice");
                        break;
                    case "Thu nợ":
                        rs = db.OpenRecordset("paid");
                        break;
                }

                if (!rs.EOF)
                    rs.MoveLast();

                if (rs.RecordCount > 0)
                {
                    if (!rs.BOF)
                        rs.MoveFirst();

                    double i = 1;
                    while (!rs.EOF)
                    {
                        switch (fieldSearch.Text)
                        {
                            case "Mã số thuế":
                                searchValue = Convert.ToString(rs.Fields["mst"].Value);
                                matchSearchCondition = searchValue.Contains(searchWhat);
                                if (matchSearchCondition)
                                {
                                    mst = rs.Fields["mst"].Value;
                                    ten_don_vi = rs.Fields["cong_ty"].Value;
                                    recordNumber = i;
                                }
                                break;
                            case "Tên đơn vị":
                                searchValue = Convert.ToString(rs.Fields["cong_ty"].Value);
                                matchSearchCondition = searchValue.Contains(searchWhat);
                                if (matchSearchCondition)
                                {
                                    mst = rs.Fields["mst"].Value;
                                    ten_don_vi = rs.Fields["cong_ty"].Value;
                                    recordNumber = i;
                                }
                                break;
                            case "Số hóa đơn":
                                if (rs.Name == "invoice")
                                {
                                    searchValue = Convert.ToString(rs.Fields["so_hoa_don"].Value);
                                    matchSearchCondition = searchValue.Contains(searchWhat);
                                    if (matchSearchCondition)
                                    {
                                        ten_don_vi = rs.Fields["mst"].Value;
                                        mst = rs.Fields["mst"].Value;
                                        ma_hoa_don = rs.Fields["ki_hieu_hoa_don"].Value;
                                        so_hoa_don = rs.Fields["so_hoa_don"].Value;
                                        so_tien = Convert.ToDouble(rs.Fields["so_tien_phat_sinh"].Value);
                                        recordNumber = i;
                                    }
                                }
                                else
                                {
                                    searchValue = Convert.ToString(rs.Fields["so_hoa_don"].Value);
                                    matchSearchCondition = searchValue.Contains(searchWhat);
                                    if (matchSearchCondition)
                                    {
                                        ma_hoa_don = rs.Fields["ki_hieu_hoa_don"].Value;
                                        so_hoa_don = rs.Fields["so_hoa_don"].Value;
                                        so_tien = Convert.ToDouble(rs.Fields["so_tien_thanh_toan"].Value);
                                        recordNumber = i;
                                    }
                                }
                                break;
                            case "Số tiền":
                                if (rs.Name == "invoice")
                                {
                                    searchDouble = Convert.ToDouble(rs.Fields["so_tien_phat_sinh"].Value);
                                    matchSearchCondition = searchDouble.Equals(Convert.ToDouble(searchWhat));
                                    if (matchSearchCondition)
                                    {
                                        mst = rs.Fields["mst"].Value;
                                        ma_hoa_don = rs.Fields["ki_hieu_hoa_don"].Value;
                                        so_hoa_don = rs.Fields["so_hoa_don"].Value;
                                        so_tien = Convert.ToDouble(rs.Fields["so_tien_phat_sinh"].Value);
                                        recordNumber = i;
                                    }
                                }
                                else
                                {
                                    searchDouble = Convert.ToDouble(rs.Fields["so_tien_thanh_toan"].Value);
                                    matchSearchCondition = searchDouble.Equals(Convert.ToDouble(searchWhat));
                                    if (matchSearchCondition)
                                    {
                                        ma_hoa_don = rs.Fields["ki_hieu_hoa_don"].Value;
                                        so_hoa_don = rs.Fields["so_hoa_don"].Value;
                                        so_tien = Convert.ToDouble(rs.Fields["so_tien_thanh_toan"].Value);
                                        recordNumber = i;
                                    }
                                }
                                break;
                            default:
                                MessageBox.Show("Lỗi tìm kiếm", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                break;
                        }

                        if (matchSearchCondition)
                            searchList.Rows.Add(mst, ten_don_vi, ma_hoa_don, so_hoa_don, so_tien, recordNumber);

                        rs.MoveNext();
                        i++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message.ToString(), "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void TbSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                search.PerformClick();

        }
    }
}
