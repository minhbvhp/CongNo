﻿using System;
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
        public TextBox[] textBoxes = new TextBox[5];
        public DateTimePicker[] dateTimePickers = new DateTimePicker[4];

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
        public void CurrentInfoRefresh(RefreshOption refreshOption)
        {
            currentMST.Clear();
            currentKhachHang.Clear();
            currentMaHoaDon.Clear();
            currentSoHoaDon.Clear();
            currentHanTra.Clear();
            currentSoTienNo.Clear();
            currentNgayChungTu.Clear();
            currentNgayHoaDon.Clear();
            currentNgayTra.Clear();
            currentSoTienTra.Clear();
            recNo.ResetText();
            modify.Visible = false;
            delete.Visible = false;

            if (refreshOption == RefreshOption.All)
            {
                modifyGroup.Visible = false;
            }
            
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            searchBy["Khách hàng"] = new string[] { "Mã số thuế", "Tên đơn vị" };
            searchBy["Phát sinh"] = new string[] { "Số hóa đơn", "Số tiền" };
            searchBy["Thu nợ"] = new string[] { "Số hóa đơn", "Số tiền" };

            foreach (String search in searchBy.Keys)
                categorySearch.Items.Add(search);

            textBoxes[0] = afterKhachHang;
            textBoxes[1] = afterMaHoaDon;
            textBoxes[2] = afterSoHoaDon;
            textBoxes[3] = afterSoTienNo;
            textBoxes[4] = afterSoTienTra;

            dateTimePickers[0] = afterHanTra;
            dateTimePickers[1] = afterNgayChungTu;
            dateTimePickers[2] = afterNgayHoaDon;
            dateTimePickers[3] = afterNgayTra;
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
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "File mẫu|Mau lay du lieu Sunweb.xlsx";
            openFileDialog.FileName = "Mau lay du lieu Sunweb";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                String inputPath = Path.GetFullPath(openFileDialog.FileName);
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
            CurrentInfoRefresh(RefreshOption.All);
            searchList.Enabled = false;
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
                        searchList.Columns["mst"].Visible = false;
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
            //Tạo mẫu báo cáo
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.DefaultExt = "xlsx";
                saveFileDialog.Filter = "Excel Workbook(*.xlsx)|*.xlsx";
                saveFileDialog.FileName = "Doi chieu cong no";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var newFile = new FileInfo(saveFileDialog.FileName);

                    if (newFile.Exists)
                        newFile.Delete();

                    using (var package = new ExcelPackage(newFile))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Doi chieu cong no");

                        worksheet.Column(1).Width = GetTrueColumnWidth(12.14);
                        worksheet.Column(2).Width = GetTrueColumnWidth(7.57);
                        worksheet.Column(3).Width = GetTrueColumnWidth(6.86);
                        worksheet.Column(4).Width = GetTrueColumnWidth(80.00);
                        worksheet.Column(5).Width = GetTrueColumnWidth(15.00);
                        worksheet.Column(6).Width = GetTrueColumnWidth(15.00);
                        worksheet.Column(7).Width = GetTrueColumnWidth(15.00);
                        worksheet.Column(8).Width = GetTrueColumnWidth(15.00);
                        worksheet.Column(9).Width = GetTrueColumnWidth(10.00);

                        for (int i = 10; i < 45; i++)
                        {
                            worksheet.Column(i).Width = GetTrueColumnWidth(20.00);
                        }

                        worksheet.Cells["A:AS"].Style.Font.Name = "Times New Roman ";
                        worksheet.Cells["A:AS"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        //Cell ngày đối chiếu
                        worksheet.Row(1).Height = 35.25;
                        worksheet.Cells["E1"].Style.Font.Bold = true;
                        worksheet.Cells["E1"].Style.Border.BorderAround(ExcelBorderStyle.Double);
                        worksheet.Cells["E1"].Style.Numberformat.Format = "dd/MM/yyyy";
                        worksheet.Cells["E1"].Value = DateTime.Today;
                        worksheet.Cells["E1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells["E1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 192, 0));
                        worksheet.Cells["E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                        /*
                        worksheet.Cells["F1"].Style.Font.Bold = true;
                        worksheet.Cells["F1"].Style.Font.Color.SetColor(Color.FromArgb(149, 55, 53));
                        

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
                        */
                        package.SaveAs(newFile);
                        MessageBox.Show("Đã lập đối chiếu công nợ", "Chúc mừng", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

            }
        }

        private void Search_Click(object sender, EventArgs e)
        {
            searchList.Enabled = true;
            searchList.Rows.Clear();
            CurrentInfoRefresh(RefreshOption.All);
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
                int rcNumber = 0;
                int i = 1;

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
                                }
                                break;
                            case "Tên đơn vị":
                                searchValue = Convert.ToString(rs.Fields["cong_ty"].Value);
                                matchSearchCondition = searchValue.Contains(searchWhat);
                                if (matchSearchCondition)
                                {
                                    mst = rs.Fields["mst"].Value;
                                    ten_don_vi = rs.Fields["cong_ty"].Value;
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
                                    }
                                }
                                break;
                            case "Số tiền":
                                if (rs.Name == "invoice")
                                {
                                    searchDouble = Convert.ToDouble(rs.Fields["so_tien_phat_sinh"].Value);
                                    matchSearchCondition = searchDouble == Convert.ToDouble(searchWhat);
                                    if (matchSearchCondition)
                                    {
                                        mst = rs.Fields["mst"].Value;
                                        ma_hoa_don = rs.Fields["ki_hieu_hoa_don"].Value;
                                        so_hoa_don = rs.Fields["so_hoa_don"].Value;
                                        so_tien = Convert.ToDouble(rs.Fields["so_tien_phat_sinh"].Value);
                                    }
                                }
                                else
                                {
                                    searchDouble = Convert.ToDouble(rs.Fields["so_tien_thanh_toan"].Value);
                                    matchSearchCondition = searchDouble == Convert.ToDouble(searchWhat);
                                    if (matchSearchCondition)
                                    {
                                        ma_hoa_don = rs.Fields["ki_hieu_hoa_don"].Value;
                                        so_hoa_don = rs.Fields["so_hoa_don"].Value;
                                        so_tien = Convert.ToDouble(rs.Fields["so_tien_thanh_toan"].Value);
                                    }
                                }
                                break;
                            default:
                                MessageBox.Show("Lỗi tìm kiếm", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                        }

                        rcNumber = i;
                        if (matchSearchCondition)
                            searchList.Rows.Add(mst, ten_don_vi, ma_hoa_don, so_hoa_don, so_tien, rcNumber);

                        rs.MoveNext();
                        i++;
                    }
                }
                rs.Close();
                db.Close();
            }
            catch
            {
                MessageBox.Show("Lỗi tìm kiếm", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SearchList_SelectionChanged(object sender, EventArgs e)
        {
            CurrentInfoRefresh(RefreshOption.InfoOnly);

            String db_name = "2019.mdb";
            String db_path = Environment.CurrentDirectory + @"\Database\";
            String db_file = db_path + db_name;

            dao.DBEngine dBEngine = new dao.DBEngine();
            dao.Database db;
            dao.Recordset rs = null;

            int recordNumber = 0;

            if (searchList.SelectedRows.Count != 0)
            {
                DataGridViewRow viewRow = searchList.SelectedRows[0];
                recordNumber = Convert.ToInt32(viewRow.Cells["recNumber"].Value) - 1;
                recNo.Text = (recordNumber + 1).ToString();
                modify.Visible = true;
            }

            try
            {
                db = dBEngine.OpenDatabase(db_file);

                String field = searchBy.Keys.ElementAt(categorySearch.SelectedIndex);

                switch (fieldSearch.Text)
                {
                    case "Mã số thuế":
                    case "Tên đơn vị":
                        rs = db.OpenRecordset("customers");
                        if (!rs.BOF)
                            rs.MoveFirst();
                        rs.Move(recordNumber);
                        currentMST.Text = rs.Fields["mst"].Value;
                        currentKhachHang.Text = rs.Fields["cong_ty"].Value;
                        break;
                    case "Số hóa đơn":
                    case "Số tiền":
                        delete.Visible = true;

                        if (field == "Phát sinh")
                        {
                            rs = db.OpenRecordset("invoice");
                            if (!rs.BOF)
                                rs.MoveFirst();
                            rs.Move(recordNumber);
                            currentMST.Text = rs.Fields["mst"].Value;
                            dao.Recordset rsKhachHang = db.OpenRecordset("SELECT cong_ty FROM customers WHERE mst = '" + currentMST.Text + "'");
                            if (rsKhachHang.RecordCount > 0)
                                currentKhachHang.Text = rsKhachHang.Fields["cong_ty"].Value;
                            currentMaHoaDon.Text = rs.Fields["ki_hieu_hoa_don"].Value;
                            currentSoHoaDon.Text = rs.Fields["so_hoa_don"].Value;
                            currentHanTra.Text = String.Format("{0:dd/MM/yyyy}", rs.Fields["han_thanh_toan"].Value);
                            currentSoTienNo.Text = String.Format("{0:n0}", rs.Fields["so_tien_phat_sinh"].Value);
                            currentNgayChungTu.Text = String.Format("{0:dd/MM/yyyy}", rs.Fields["ngay_ct"].Value);
                            currentNgayHoaDon.Text = String.Format("{0:dd/MM/yyyy}", rs.Fields["ngay_hoa_don"].Value);
                        }
                        else if (field == "Thu nợ")
                        {
                            rs = db.OpenRecordset("paid");
                            if (!rs.BOF)
                                rs.MoveFirst();
                            rs.Move(recordNumber);
                            currentMaHoaDon.Text = rs.Fields["ki_hieu_hoa_don"].Value;
                            currentSoHoaDon.Text = rs.Fields["so_hoa_don"].Value;
                            currentNgayTra.Text = String.Format("{0:dd/MM/yyyy}", rs.Fields["ngay_thanh_toan"].Value);
                            currentSoTienTra.Text = String.Format("{0:n0}", rs.Fields["so_tien_thanh_toan"].Value);
                        }
                        break;
                }
                rs.Close();
                db.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message.ToString(), "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //Export file "Mau lay du lieu Sunweb.xlsx"
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
                        worksheet.Cells["A:O"].Style.Font.Name = "Calibri";
                        worksheet.Cells["A1:O1"].Style.Font.Bold = true;

                        worksheet.Column(1).Width = GetTrueColumnWidth(6);
                        worksheet.Column(2).Width = GetTrueColumnWidth(9);
                        worksheet.Column(3).Width = GetTrueColumnWidth(12);
                        worksheet.Column(4).Width = GetTrueColumnWidth(5);
                        worksheet.Column(5).Width = GetTrueColumnWidth(12);
                        worksheet.Column(6).Width = GetTrueColumnWidth(35);
                        worksheet.Column(7).Width = GetTrueColumnWidth(30);
                        worksheet.Column(8).Width = GetTrueColumnWidth(11);
                        worksheet.Column(9).Width = GetTrueColumnWidth(8);
                        worksheet.Column(10).Width = GetTrueColumnWidth(8.5);
                        worksheet.Column(11).Width = GetTrueColumnWidth(13);
                        worksheet.Column(12).Width = GetTrueColumnWidth(13);
                        worksheet.Column(13).Width = GetTrueColumnWidth(9);
                        worksheet.Column(14).Width = GetTrueColumnWidth(15);
                        worksheet.Column(15).Width = GetTrueColumnWidth(15.20);

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

        private void Modify_Click(object sender, EventArgs e)
        {
            modifyGroup.Visible = true;
            searchList.Enabled = false;

            afterKhachHang.Text = currentKhachHang.Text;
            afterMaHoaDon.Text = currentMaHoaDon.Text;
            afterSoHoaDon.Text = currentSoHoaDon.Text;
            afterHanTra.Text = currentHanTra.Text;
            afterSoTienNo.Text = currentSoTienNo.Text;
            afterNgayChungTu.Text = currentNgayChungTu.Text;
            afterNgayHoaDon.Text = currentNgayHoaDon.Text;
            afterNgayTra.Text = currentNgayTra.Text;
            afterSoTienTra.Text = currentSoTienTra.Text;
        }

        private void Delete_Click(object sender, EventArgs e)
        {
            String db_name = "2019.mdb";
            String db_path = Environment.CurrentDirectory + @"\Database\";
            String db_file = db_path + db_name;

            dao.DBEngine dBEngine = new dao.DBEngine();
            dao.Database db;
            dao.Recordset rs = null;
            int recordNumber = Convert.ToInt32(recNo.Text) - 1;

            try
            {
                db = dBEngine.OpenDatabase(db_file);

                String field = searchBy.Keys.ElementAt(categorySearch.SelectedIndex);

                switch (fieldSearch.Text)
                {
                    case "Mã số thuế":
                    case "Tên đơn vị":
                        rs = db.OpenRecordset("customers");
                        break;
                    case "Số hóa đơn":
                    case "Số tiền":
                        if (field == "Phát sinh")
                        {
                            rs = db.OpenRecordset("invoice");
                        }
                        else if (field == "Thu nợ")
                        {
                            rs = db.OpenRecordset("paid");
                        }
                        break;
                }

                DialogResult dialogResult = MessageBox.Show("Bạn có chắc chắn muốn xóa thông tin này không?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Yes)
                {
                    if (!rs.BOF)
                        rs.MoveFirst();
                    rs.Move(recordNumber);
                    rs.Delete();
                }

                rs.Close();
                db.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message.ToString(), "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                searchList.Rows.Clear();
                CurrentInfoRefresh(RefreshOption.All);
            }
        }

        private void Accept_Click(object sender, EventArgs e)
        {
            int emptyAmount = 0;

            foreach (TextBox textBox in textBoxes)
            {
                if (textBox.Enabled == true)
                {
                    if (String.IsNullOrEmpty(textBox.Text.Trim()))
                    {
                        textBox.BackColor = Color.Red;
                        textBox.Refresh();
                        emptyAmount++;
                    }
                }
            }

            if (emptyAmount > 0)
            {
                MessageBox.Show("Thông tin không thể để trống");
            }
            else
            {
                String db_name = "2019.mdb";
                String db_path = Environment.CurrentDirectory + @"\Database\";
                String db_file = db_path + db_name;

                dao.DBEngine dBEngine = new dao.DBEngine();
                dao.Database db;
                dao.Recordset rs = null;
                int recordNumber = Convert.ToInt32(recNo.Text) - 1;

                try
                {
                    db = dBEngine.OpenDatabase(db_file);

                    String field = searchBy.Keys.ElementAt(categorySearch.SelectedIndex);

                    switch (fieldSearch.Text)
                    {
                        case "Mã số thuế":
                        case "Tên đơn vị":
                            rs = db.OpenRecordset("customers");
                            if (!rs.BOF)
                                rs.MoveFirst();
                            rs.Move(recordNumber);
                            rs.Edit();
                            rs.Fields["cong_ty"].Value = afterKhachHang.Text;
                            break;
                        case "Số hóa đơn":
                        case "Số tiền":
                            if (field == "Phát sinh")
                            {
                                rs = db.OpenRecordset("invoice");
                                if (!rs.BOF)
                                    rs.MoveFirst();
                                rs.Move(recordNumber);
                                rs.Edit();
                                rs.Fields["ki_hieu_hoa_don"].Value = afterMaHoaDon.Text;
                                rs.Fields["so_hoa_don"].Value = afterSoHoaDon.Text;
                                rs.Fields["han_thanh_toan"].Value = afterHanTra.Text;
                                rs.Fields["so_tien_phat_sinh"].Value = afterSoTienNo.Text;
                                rs.Fields["ngay_ct"].Value = afterNgayChungTu.Text;
                                rs.Fields["ngay_hoa_don"].Value = afterNgayHoaDon.Text;
                            }
                            else if (field == "Thu nợ")
                            {
                                rs = db.OpenRecordset("paid");
                                if (!rs.BOF)
                                    rs.MoveFirst();
                                rs.Move(recordNumber);
                                rs.Edit();
                                rs.Fields["ki_hieu_hoa_don"].Value = afterMaHoaDon.Text;
                                rs.Fields["so_hoa_don"].Value = afterSoHoaDon.Text;
                                rs.Fields["ngay_thanh_toan"].Value = afterNgayTra.Text;
                                rs.Fields["so_tien_thanh_toan"].Value = afterSoTienTra.Text;
                            }
                            break;
                    }

                    DialogResult dialogResult = MessageBox.Show("Bạn có chắc chắn muốn chỉnh sửa thông tin này không?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dialogResult == DialogResult.Yes)
                    {
                        rs.Update();
                    }
                    else
                    {
                        rs.CancelUpdate();
                    }
                    rs.Close();
                    db.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message.ToString(), "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    searchList.Rows.Clear();
                    CurrentInfoRefresh(RefreshOption.All);
                }
            }
            searchList.Rows.Clear();
            CurrentInfoRefresh(RefreshOption.All);
        }

        private void FieldSearch_SelectedIndexChanged(object sender, EventArgs e)
        {
            searchList.Enabled = false;
            CurrentInfoRefresh(RefreshOption.All);
        }

        private void ModifyGroup_VisibleChanged(object sender, EventArgs e)
        {
            afterKhachHang.Clear();
            afterMaHoaDon.Clear();
            afterSoHoaDon.Clear();
            afterHanTra.ResetText();
            afterSoTienTra.Clear();
            afterNgayChungTu.ResetText();
            afterNgayHoaDon.ResetText();
            afterNgayTra.ResetText();
            afterSoTienTra.Clear();

            if (modifyGroup.Visible == true)
            {
                foreach (TextBox textBox in textBoxes)
                {
                    textBox.BackColor = Control.DefaultBackColor;
                    textBox.Enabled = false;
                }

                foreach (DateTimePicker dateTimePicker in dateTimePickers)
                    dateTimePicker.Enabled = false;

                String field = searchBy.Keys.ElementAt(categorySearch.SelectedIndex);

                switch (fieldSearch.Text)
                {
                    case "Mã số thuế":
                    case "Tên đơn vị":
                        afterKhachHang.Enabled = true;
                        afterKhachHang.BackColor = SystemColors.Window;
                        break;
                    case "Số hóa đơn":
                    case "Số tiền":
                        afterMaHoaDon.Enabled = true;
                        afterMaHoaDon.BackColor = SystemColors.Window;

                        afterSoHoaDon.Enabled = true;
                        afterSoHoaDon.BackColor = SystemColors.Window;

                        if (field == "Phát sinh")
                        {
                            afterHanTra.Enabled = true;
                            afterHanTra.BackColor = SystemColors.Window;

                            afterSoTienNo.Enabled = true;
                            afterSoTienNo.BackColor = SystemColors.Window;

                            afterNgayChungTu.Enabled = true;
                            afterNgayChungTu.BackColor = SystemColors.Window;

                            afterNgayHoaDon.Enabled = true;
                            afterNgayHoaDon.BackColor = SystemColors.Window;
                        }
                        else if (field == "Thu nợ")
                        {
                            afterNgayTra.Enabled = true;
                            afterNgayTra.BackColor = SystemColors.Window;

                            afterSoTienTra.Enabled = true;
                            afterSoTienTra.BackColor = SystemColors.Window;
                        }
                        break;
                    default:
                        return;
                }
            }
        }

        public static double GetTrueColumnWidth(double width)
        {
            //DEDUCE WHAT THE COLUMN WIDTH WOULD REALLY GET SET TO
            double z = 1d;
            if (width >= (1 + 2 / 3))
            {
                z = Math.Round((Math.Round(7 * (width - 1 / 256), 0) - 5) / 7, 2);
            }
            else
            {
                z = Math.Round((Math.Round(12 * (width - 1 / 256), 0) - Math.Round(5 * width, 0)) / 12, 2);
            }

            //HOW FAR OFF? (WILL BE LESS THAN 1)
            double errorAmt = width - z;

            //CALCULATE WHAT AMOUNT TO TACK ONTO THE ORIGINAL AMOUNT TO RESULT IN THE CLOSEST POSSIBLE SETTING 
            double adj = 0d;
            if (width >= (1 + 2 / 3))
            {
                adj = (Math.Round(7 * errorAmt - 7 / 256, 0)) / 7;
            }
            else
            {
                adj = ((Math.Round(12 * errorAmt - 12 / 256, 0)) / 12) + (2 / 12);
            }

            //RETURN A SCALED-VALUE THAT SHOULD RESULT IN THE NEAREST POSSIBLE VALUE TO THE TRUE DESIRED SETTING
            if (z > 0)
            {
                return width + adj;
            }

            return 0d;
        }
    }
}