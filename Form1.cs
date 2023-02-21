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
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Access.Dao;

namespace CongNo
{
    
    public partial class Form1 : Form
    {
        
        public TextBox[] textBoxes = new TextBox[6];
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
            currentNgayPhatSinh.Clear();
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

        public int DateToColumnTienVe(DateTime dateTime)
        {
            int month = dateTime.Month;
            return month;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = "Số liệu năm " + Program.DbYear;
            loaiTk.SelectedItem = "Tổng";
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
            textBoxes[5] = afterMST;

            dateTimePickers[0] = afterHanTra;
            dateTimePickers[1] = afterNgayPhatSinh;
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

            //Read Sunweb excel form
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "File mẫu|*.xlsx";
            openFileDialog.FileName = "Mau lay du lieu Sunweb";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                String inputPath = Path.GetFullPath(openFileDialog.FileName);
                var importExcel = new FileInfo(inputPath);
                using (var package = new ExcelPackage(importExcel))
                {
                    //Get total of records
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    int lastRow = (int)worksheet.Dimension.End.Row;

                    //Create connection
                    int row;

                    String db_name = Program.DbYear + ".mdb";
                    String db_path = Environment.CurrentDirectory + @"\Database\";
                    String db_file = db_path + db_name;

                    DBEngine dBEngine = new DBEngine();
                    Database db = null;
                    Recordset rs;

                    //Export to .mdb file
                    try
                    {
                        db = dBEngine.OpenDatabase(db_file);
                        db.BeginTrans();
                        db.Execute("draft_clear");
                        rs = db.OpenRecordset("draft");
                        for (row = 2; row <= lastRow; row++)
                        {
                            rs.AddNew();
                            rs.Fields["dong"].Value = row;
                            rs.Fields["loai_ct"].Value = NullToString(worksheet.Cells["A" + row].Value);
                            rs.Fields["no_co"].Value = NullToString(worksheet.Cells["E" + row].Value);
                            rs.Fields["mst_draft"].Value = NullToString(worksheet.Cells["P" + row].Value);
                            rs.Fields["cong_ty1"].Value = NullToString(worksheet.Cells["G" + row].Value);
                            rs.Fields["cong_ty2"].Value = NullToString(worksheet.Cells["H" + row].Value);
                            rs.Fields["ky_hieu_hd"].Value = NullToString(worksheet.Cells["Q" + row].Value);
                            rs.Fields["so_hoa_don"].Value = NullToString(worksheet.Cells["R" + row].Value);
                            rs.Fields["ngay_hoa_don_draft"].Value = NullToString(worksheet.Cells["S" + row].Value);
                            rs.Fields["ma_nv"].Value = NullToString(worksheet.Cells["L" + row].Value);
                            rs.Fields["ma_phong"].Value = NullToString(worksheet.Cells["M" + row].Value);
                            rs.Fields["so_tien"].Value = NullToString(worksheet.Cells["D" + row].Value);
                            rs.Fields["han_tt_draft"].Value = NullToString(worksheet.Cells["I" + row].Value);
                            rs.Fields["ngay_ct_draft"].Value = NullToString(worksheet.Cells["B" + row].Value);
                            rs.Fields["user"].Value = NullToString(worksheet.Cells["T" + row].Value);
                            rs.Fields["kenh_kt"].Value = NullToString(worksheet.Cells["N" + row].Value);
                            rs.Fields["so_tai_khoan"].Value = NullToString(worksheet.Cells["C" + row].Value);
                            rs.Fields["so_tham_chieu"].Value = NullToString(worksheet.Cells["F" + row].Value);
                            rs.Fields["loai_tien_draft"].Value = NullToString(worksheet.Cells["J" + row].Value);

                            string _tempMaCanBo = NullToString(worksheet.Cells["O" + row].Value) ?? "";
                            if (String.IsNullOrEmpty(_tempMaCanBo) || String.IsNullOrWhiteSpace(_tempMaCanBo)){
                                rs.Fields["ma_can_bo_draft"].Value = "CB1460000";
                            }
                            else
                            {
                                rs.Fields["ma_can_bo_draft"].Value = NullToString(worksheet.Cells["O" + row].Value);
                            }

                            if (rs.Fields["loai_tien_draft"].Value == "USD")
                            {
                            rs.Fields["nguyen_te_draft"].Value = NullToString(worksheet.Cells["K" + row].Value);
                            }

                            rs.Update();

                            uploadProgress.Value = (row - 1) * 100 / lastRow;
                            uploadProgress.Refresh();
                            Application.DoEvents();
                        }

                        db.Execute("update_mst");
                        db.Execute("add_mst_to_customers");

                        db.Execute("invoice_filter");
                        db.Execute("add_invoice");

                        db.Execute("paid_filter");
                        db.Execute("add_paid");

                        db.Execute("update_revenue");
                        db.Execute("add_revenue");

                        uploadProgress.Value += 1;
                        uploadProgress.Refresh();

                        //Show not_upload to DataGridView
                        rs = db.OpenRecordset("not_upload");
                        if (!rs.EOF)
                            rs.MoveLast();

                        if (rs.RecordCount > 0)
                        {
                            String dong;
                            String ten_cong_ty;
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
                                ky_hieu_hd = Convert.ToString(rs.Fields["ky_hieu_hd"].Value);
                                so_hoa_don = Convert.ToString(rs.Fields["so_hoa_don"].Value);
                                so_tien = rs.Fields["so_tien"].Value;
                                user = Convert.ToString(rs.Fields["user"].Value);
                                ghiChu = GhiChu(so_hoa_don, ten_cong_ty, ky_hieu_hd);

                                notUploadList.Rows.Add(dong, ten_cong_ty, ky_hieu_hd, so_hoa_don, so_tien, user, ghiChu);
                                rs.MoveNext();
                            }
                        }

                        db.Execute("draft_clear");
                        db.Execute("invoice_draft_clear");
                        db.Execute("paid_draft_clear");
                        db.CommitTrans();

                        rs.Close();                        

                        MessageBox.Show("Đã upload dữ liệu xong.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Không ghi được dữ liệu.\n" + ex.Message.ToString(), "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        db.Rollback();
                    }
                    finally
                    {
                        db.Close();                        
                    }                    
                }
            }
            CurrentInfoRefresh(RefreshOption.All);
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
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.DefaultExt = "xlsm";
                    saveFileDialog.Filter = "Excel Workbook(*.xlsm)|*.xlsm";
                    saveFileDialog.FileName = "Doi chieu cong no - " + Program.DbYear + " (" + loaiTk.Text + ")";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        int i;
                        String db_name = Program.DbYear + ".mdb";
                        String db_path = Environment.CurrentDirectory + @"\Database\";
                        String db_file = db_path + db_name;

                        DBEngine dBEngine = new DBEngine();
                        Database db;
                        Recordset rs;

                        db = dBEngine.OpenDatabase(db_file);

                        //Get department list from Database
                        Dictionary<String, String> departments = new Dictionary<String, String>();
                        rs = db.OpenRecordset("department");
                        if (!rs.BOF)
                            rs.MoveFirst();
                        for (i = 1; i <= rs.RecordCount; i++)
                        {
                            departments.Add(rs.Fields["ten_phong"].Value, rs.Fields["ten_day_du"].Value);
                            rs.MoveNext();
                        }                        

                        //Create "Doi chieu cong no" form
                        var newFile = new FileInfo(saveFileDialog.FileName);

                        if (newFile.Exists)
                            newFile.Delete();

                        using (var package = new ExcelPackage(newFile))
                        {
                            package.Workbook.Properties.Title = "Doi chieu cong no - " + Program.DbYear;
                            package.Workbook.Properties.Company = "Bảo Việt Hải Phòng";

                            //Sheet department list
                            ExcelWorksheet departmentWorksheet = package.Workbook.Worksheets.Add("List Phong");
                            int rowOfDepartmentWorksheet = 1;

                            foreach (KeyValuePair<String, String> department in departments)
                            {
                                departmentWorksheet.Cells["A" + rowOfDepartmentWorksheet.ToString()].Value = department.Key;
                                departmentWorksheet.Cells["B" + rowOfDepartmentWorksheet.ToString()].Value = department.Value;
                                rowOfDepartmentWorksheet++;
                            }

                            departmentWorksheet.Hidden = eWorkSheetHidden.VeryHidden;

                            //Sheet "Doi chieu cong no"
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

                            for (i = 10; i <= 45; i++)
                            {
                                worksheet.Column(i).Width = GetTrueColumnWidth(20.00);
                            }

                            for (i = 46; i <= 50; i++)
                            {
                                worksheet.Column(i).Width = GetTrueColumnWidth(10.00);
                            }

                            worksheet.Cells["A:AY"].Style.Font.Name = "Times New Roman";
                            worksheet.Cells["A:AY"].Style.Font.Size = 11;
                            worksheet.Cells["A:AY"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Row(1).Height = 35.25;
                            worksheet.Row(2).Height = 9.75;
                            worksheet.Row(3).Height = 35.25;
                            worksheet.Row(4).Height = 35.25;
                            worksheet.Row(5).Height = 25.5;
                            worksheet.Row(6).Height = 35.25;
                            worksheet.Row(7).Height = 38.25;
                            worksheet.Row(8).Height = 15.75;

                            //Cell ngày đối chiếu
                            worksheet.Cells["E1"].Style.Font.Bold = true;
                            worksheet.Cells["E1"].Style.Font.Size = 12;
                            worksheet.Cells["E1"].Style.Border.BorderAround(ExcelBorderStyle.Double);
                            worksheet.Cells["E1"].Style.Numberformat.Format = "dd/MM/yyyy";
                            worksheet.Cells["E1"].Value = DateTime.Today;
                            worksheet.Cells["E1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["E1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 192, 0));
                            worksheet.Cells["E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            //Cell tên phòng
                            worksheet.Cells["F1"].Style.Font.Bold = true;
                            worksheet.Cells["F1"].Style.Font.Size = 14;
                            worksheet.Cells["F1"].Style.Border.BorderAround(ExcelBorderStyle.Double);
                            worksheet.Cells["F1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["F1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(250, 192, 144));
                            worksheet.Cells["F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            var val = worksheet.Cells["F1"].DataValidation.AddListDataValidation();
                            val.Formula.Values.Add("Công ty");
                            foreach (String department in departments.Keys)
                                val.Formula.Values.Add(department);
                            worksheet.Cells["F1"].Value = "Công ty";

                            //Cell Phòng
                            worksheet.Cells["D3"].Formula = @"IF(F1<>""Công ty"",""PHÒNG "" & VLOOKUP(F1,'List Phong'!A1:B12,2,FALSE), ""BẢO VIỆT HẢI PHÒNG"")";
                            worksheet.Cells["D3"].Style.Font.Bold = true;
                            worksheet.Cells["D3"].Style.Font.Size = 12;

                            //Cell Title
                            worksheet.Cells["C4"].Value = "BẢNG ĐỐI CHIẾU NỢ PHẢI THU CHI TIẾT THEO TỪNG KHÁCH HÀNG";
                            worksheet.Cells["C4:AS4"].Merge = true;
                            worksheet.Cells["C4:AS4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells["C4:AS4"].Style.Font.Bold = true;
                            worksheet.Cells["C4:AS4"].Style.Font.Size = 14;

                            worksheet.Cells["C5"].Formula = @"""Tháng "" & MONTH(E1) & "" năm "" & YEAR(E1)";
                            worksheet.Cells["C5:AS5"].Merge = true;
                            worksheet.Cells["C5:AS5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells["C5:AS5"].Style.Font.Bold = true;
                            worksheet.Cells["C5:AS5"].Style.Font.Size = 14;

                            //Column name
                            worksheet.Cells["A6:AS8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["A6:B8"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 0));
                            worksheet.Cells["C6:D8"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 217, 195));
                            worksheet.Cells["E6:H8"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 192, 218));
                            worksheet.Cells["I6:J8"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 217, 195));
                            worksheet.Cells["K6:W8"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(230, 185, 184));
                            worksheet.Cells["X6:AJ8"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(197, 217, 241));
                            worksheet.Cells["AK6:AS8"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(221, 217, 195));

                            worksheet.Cells["A6:A7"].Merge = true;
                            worksheet.Cells["A6:A7"].Value = "Ngày hóa đơn";

                            worksheet.Cells["B6:B7"].Merge = true;
                            worksheet.Cells["B6:B7"].Value = "Phòng";

                            worksheet.Cells["C6:C7"].Merge = true;
                            worksheet.Cells["C6:C7"].Value = "STT";

                            worksheet.Cells["D6:D7"].Merge = true;
                            worksheet.Cells["D6:D7"].Value = "Tên đơn vị";

                            worksheet.Cells["E6:H6"].Merge = true;
                            worksheet.Cells["E6:H6"].Value = "Thông tin";
                            worksheet.Cells["E7"].Value = "Mã hóa đơn";
                            worksheet.Cells["F7"].Value = "Hóa đơn";
                            worksheet.Cells["G7"].Value = "Hạn thanh toán";
                            worksheet.Cells["H7"].Value = "Nghiệp vụ";

                            worksheet.Cells["I6:I7"].Merge = true;
                            worksheet.Cells["I6:I7"].Value = "Số ngày quá hạn";

                            worksheet.Cells["J6:J7"].Merge = true;
                            worksheet.Cells["J6:J7"].Value = "Dư đầu kỳ";

                            worksheet.Cells["K6:W6"].Merge = true;
                            worksheet.Cells["K6:W6"].Value = "Phát sinh tháng nợ";
                            worksheet.Cells["K7"].Value = "Tháng 01";
                            worksheet.Cells["L7"].Value = "Tháng 02";
                            worksheet.Cells["M7"].Value = "Tháng 03";
                            worksheet.Cells["N7"].Value = "Tháng 04";
                            worksheet.Cells["O7"].Value = "Tháng 05";
                            worksheet.Cells["P7"].Value = "Tháng 06";
                            worksheet.Cells["Q7"].Value = "Tháng 07";
                            worksheet.Cells["R7"].Value = "Tháng 08";
                            worksheet.Cells["S7"].Value = "Tháng 09";
                            worksheet.Cells["T7"].Value = "Tháng 10";
                            worksheet.Cells["U7"].Value = "Tháng 11";
                            worksheet.Cells["V7"].Value = "Tháng 12";
                            worksheet.Cells["W7"].Value = "Cộng phát sinh";

                            worksheet.Cells["X6:AJ6"].Merge = true;
                            worksheet.Cells["X6:AJ6"].Value = "Theo dõi thu nợ";
                            worksheet.Cells["X7"].Value = "Tháng 01";
                            worksheet.Cells["Y7"].Value = "Tháng 02";
                            worksheet.Cells["Z7"].Value = "Tháng 03";
                            worksheet.Cells["AA7"].Value = "Tháng 04";
                            worksheet.Cells["AB7"].Value = "Tháng 05";
                            worksheet.Cells["AC7"].Value = "Tháng 06";
                            worksheet.Cells["AD7"].Value = "Tháng 07";
                            worksheet.Cells["AE7"].Value = "Tháng 08";
                            worksheet.Cells["AF7"].Value = "Tháng 09";
                            worksheet.Cells["AG7"].Value = "Tháng 10";
                            worksheet.Cells["AH7"].Value = "Tháng 11";
                            worksheet.Cells["AI7"].Value = "Tháng 12";
                            worksheet.Cells["AJ7"].Value = "Cộng thanh toán";

                            worksheet.Cells["AK6:AK7"].Merge = true;
                            worksheet.Cells["AK6:AK7"].Value = "Cuối kì";

                            worksheet.Cells["AL6:AL7"].Merge = true;
                            worksheet.Cells["AL6:AL7"].Value = "Trong hạn thanh toán";

                            worksheet.Cells["AM6:AM7"].Merge = true;
                            worksheet.Cells["AM6:AM7"].Value = "Quá hạn thanh toán dưới 1 tháng";

                            worksheet.Cells["AN6:AN7"].Merge = true;
                            worksheet.Cells["AN6:AN7"].Value = "Quá hạn thanh toán dưới 3 tháng";

                            worksheet.Cells["AO6:AO7"].Merge = true;
                            worksheet.Cells["AO6:AO7"].Value = "Quá hạn thanh toán từ 3 - 6 tháng";

                            worksheet.Cells["AP6:AP7"].Merge = true;
                            worksheet.Cells["AP6:AP7"].Value = "Quá hạn thanh toán từ 6 tháng - dưới 1 năm";

                            worksheet.Cells["AQ6:AQ7"].Merge = true;
                            worksheet.Cells["AQ6:AQ7"].Value = "Quá hạn thanh toán từ 1 - 2 năm";

                            worksheet.Cells["AR6:AR7"].Merge = true;
                            worksheet.Cells["AR6:AR7"].Value = "Quá hạn thanh toán từ 2 - 3 năm";

                            worksheet.Cells["AS6:AS7"].Merge = true;
                            worksheet.Cells["AS6:AS7"].Value = "Quá hạn thanh toán trên 3 năm";

                            worksheet.Cells["AT6:AT7"].Merge = true;
                            worksheet.Cells["AT6:AT7"].Value = "Kênh khai thác";

                            worksheet.Cells["AU6:AU7"].Merge = true;
                            worksheet.Cells["AU6:AU7"].Value = "Số tài khoản";

                            worksheet.Cells["AV6:AV7"].Merge = true;
                            worksheet.Cells["AV6:AV7"].Value = "Số bảng kê";

                            worksheet.Cells["AW6:AW7"].Merge = true;
                            worksheet.Cells["AW6:AW7"].Value = "Loại tiền";

                            worksheet.Cells["AX6:AX7"].Merge = true;
                            worksheet.Cells["AX6:AX7"].Value = "Nguyên tệ";

                            worksheet.Cells["AY6:AY7"].Merge = true;
                            worksheet.Cells["AY6:AY7"].Value = "Cán bộ";

                            worksheet.Cells["A6:AY8"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A6:AY8"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A6:AY8"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A6:AY8"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A6:AY8"].Style.Font.Bold = true;
                            worksheet.Cells["A6:AY8"].Style.WrapText = true;
                            worksheet.Cells["A6:AY8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                            //Export "cong_no" to "Doi chieu cong no" excel file                            
                            switch (loaiTk.Text)
                            {
                                case "Tổng":
                                    rs = db.OpenRecordset("cong_no");
                                    break;
                                case "13111":
                                    rs = db.OpenRecordset("SELECT * FROM cong_no WHERE m_so_tai_khoan='13111'");
                                    rs.MoveLast();
                                    break;
                                case "13161":
                                    rs = db.OpenRecordset("SELECT * FROM cong_no WHERE m_so_tai_khoan='13161'");
                                    rs.MoveLast();
                                    break;
                            }
                            const byte ROW_BEFORE_START_EXCEL = 8;
                            int maxPhatSinh = rs.RecordCount;
                            int maxRowExcel = maxPhatSinh + ROW_BEFORE_START_EXCEL;
                            int currentRow = 0;
                            int rowTong = maxRowExcel + 1;
                            rs.MoveFirst();

                            //Format cells (inclue total rows)
                            worksheet.Cells["A" + ROW_BEFORE_START_EXCEL + ":A" + rowTong].Style.Numberformat.Format = "dd/MM/yyyy";
                            worksheet.Cells["B" + ROW_BEFORE_START_EXCEL + ":B" + rowTong].Style.Numberformat.Format = "@";
                            worksheet.Cells["C" + ROW_BEFORE_START_EXCEL + ":C" + rowTong].Style.Numberformat.Format = "#";
                            worksheet.Cells["D" + ROW_BEFORE_START_EXCEL + ":D" + rowTong].Style.Numberformat.Format = "@";
                            worksheet.Cells["E" + ROW_BEFORE_START_EXCEL + ":F" + rowTong].Style.Numberformat.Format = "@";
                            worksheet.Cells["G" + ROW_BEFORE_START_EXCEL + ":G" + rowTong].Style.Numberformat.Format = "dd/MM/yyyy";
                            worksheet.Cells["H" + ROW_BEFORE_START_EXCEL + ":H" + rowTong].Style.Numberformat.Format = "@";
                            worksheet.Cells["I" + ROW_BEFORE_START_EXCEL + ":I" + rowTong].Style.Numberformat.Format = "#";
                            worksheet.Cells["J" + ROW_BEFORE_START_EXCEL + ":AS" + rowTong].Style.Numberformat.Format = "_(* #,##0_);_(* (#,##0);_(* \" - \"_);_(@_)";
                            worksheet.Cells["AX" + ROW_BEFORE_START_EXCEL + ":AX" + rowTong].Style.Numberformat.Format = "_(* #,##0.00_);_(* (#,##0.00);_(* \" - \"_);_(@_)";

                            Dictionary<String, int> invoices = new Dictionary<string, int>();
                            String invoiceID = String.Empty;
                            int invoiceRow = 0;

                            for (i = 1; i <= maxPhatSinh; i++)
                            {
                                currentRow = i + ROW_BEFORE_START_EXCEL;
                                worksheet.Cells["A" + currentRow].Value = rs.Fields["m_ngay_hoa_don"].Value;
                                worksheet.Cells["B" + currentRow].Value = rs.Fields["m_ten_phong"].Value;

                                String fSTT = String.Format("(SUBTOTAL(3,$D${0}:D{1}))", ROW_BEFORE_START_EXCEL + 1, currentRow);
                                worksheet.Cells["C" + currentRow].Formula = fSTT;

                                worksheet.Cells["D" + currentRow].Value = rs.Fields["m_cong_ty"].Value;
                                worksheet.Cells["E" + currentRow].Value = rs.Fields["ki_hieu_hoa_don"].Value;
                                worksheet.Cells["F" + currentRow].Value = rs.Fields["so_hoa_don"].Value;
                                worksheet.Cells["G" + currentRow].Value = rs.Fields["m_han_thanh_toan"].Value;
                                worksheet.Cells["H" + currentRow].Value = rs.Fields["m_ma_nv"].Value;
                                worksheet.Cells["AT" + currentRow].Value = rs.Fields["m_kenh_kt"].Value;
                                worksheet.Cells["AU" + currentRow].Value = rs.Fields["m_so_tai_khoan"].Value;
                                worksheet.Cells["AV" + currentRow].Value = rs.Fields["m_so_tham_chieu"].Value;
                                worksheet.Cells["AW" + currentRow].Value = rs.Fields["m_loai_tien"].Value;
                                worksheet.Cells["AX" + currentRow].Value = rs.Fields["m_tong_nguyen_te"].Value;
                                worksheet.Cells["AY" + currentRow].Value = rs.Fields["m_full_name"].Value;

                                String fNgayQuaHan = String.Format("IF(AND(AK{0} > 0, $E$1 >= G{0}), ($E$1 - G{0}) + 1, 0)", currentRow);
                                worksheet.Cells["I" + currentRow].Formula = fNgayQuaHan;

                                if (rs.Fields["tong_dau_ky"].Value is DBNull)
                                    worksheet.Cells["J" + currentRow].Value = 0;
                                else
                                    worksheet.Cells["J" + currentRow].Value = rs.Fields["tong_dau_ky"].Value;

                                worksheet.Cells["K" + currentRow].Value = rs.Fields["tongno1"].Value;
                                worksheet.Cells["L" + currentRow].Value = rs.Fields["tongno2"].Value;
                                worksheet.Cells["M" + currentRow].Value = rs.Fields["tongno3"].Value;
                                worksheet.Cells["N" + currentRow].Value = rs.Fields["tongno4"].Value;
                                worksheet.Cells["O" + currentRow].Value = rs.Fields["tongno5"].Value;
                                worksheet.Cells["P" + currentRow].Value = rs.Fields["tongno6"].Value;
                                worksheet.Cells["Q" + currentRow].Value = rs.Fields["tongno7"].Value;
                                worksheet.Cells["R" + currentRow].Value = rs.Fields["tongno8"].Value;
                                worksheet.Cells["S" + currentRow].Value = rs.Fields["tongno9"].Value;
                                worksheet.Cells["T" + currentRow].Value = rs.Fields["tongno10"].Value;
                                worksheet.Cells["U" + currentRow].Value = rs.Fields["tongno11"].Value;
                                worksheet.Cells["V" + currentRow].Value = rs.Fields["tongno12"].Value;

                                String fCongPhatSinh = String.Format("SUM(K{0}:V{0})", currentRow);
                                worksheet.Cells["W" + currentRow].Formula = fCongPhatSinh;

                                String fCongThanhToan = String.Format("SUM(X{0}:AI{0})", currentRow);
                                worksheet.Cells["AJ" + currentRow].Formula = fCongThanhToan;

                                String fCuoiKy = String.Format("J{0}+W{0}-AJ{0}", currentRow);
                                worksheet.Cells["AK" + currentRow].Formula = fCuoiKy;

                                String fTrongHan = String.Format("IF(I{0}=0,AK{0},0)", currentRow);
                                worksheet.Cells["AL" + currentRow].Formula = fTrongHan;

                                String fDuoi1Thang = String.Format("IF(AND(I{0}>=1,I{0}<=30),AK{0},0)", currentRow);
                                worksheet.Cells["AM" + currentRow].Formula = fDuoi1Thang;

                                String fDuoi3Thang = String.Format("IF(AND(I{0}>=31,I{0}<=90),AK{0},0)", currentRow);
                                worksheet.Cells["AN" + currentRow].Formula = fDuoi3Thang;

                                String f3Den6Thang = String.Format("IF(AND(I{0}>=91,I{0}<=180),AK{0},0)", currentRow);
                                worksheet.Cells["AO" + currentRow].Formula = f3Den6Thang;

                                String f6ThangDen1Nam = String.Format("IF(AND(I{0}>=181,I{0}<=365),AK{0},0)", currentRow);
                                worksheet.Cells["AP" + currentRow].Formula = f6ThangDen1Nam;

                                String f1Den2Nam = String.Format("IF(AND(I{0}>=366,I{0}<=730),AK{0},0)", currentRow);
                                worksheet.Cells["AQ" + currentRow].Formula = f1Den2Nam;

                                String f2Den3Nam = String.Format("IF(AND(I{0}>=731,I{0}<=1095),AK{0},0)", currentRow);
                                worksheet.Cells["AR" + currentRow].Formula = f2Den3Nam;

                                String fTren3Nam = String.Format("IF(I{0}>=1096,AK{0},0)", currentRow);
                                worksheet.Cells["AS" + currentRow].Formula = fTren3Nam;

                                //Get "ki_hieu_hoa_don" and "so_hoa_don" from "cong_no" to Dictionary
                                invoiceID = String.Format("{0};{1}", rs.Fields["ki_hieu_hoa_don"].Value, rs.Fields["so_hoa_don"].Value);
                                invoiceRow = currentRow;

                                if (!invoices.ContainsKey(invoiceID))
                                    invoices.Add(invoiceID, invoiceRow);

                                rs.MoveNext();
                            }

                            //Autofit column Ten can bo                            
                            worksheet.Column(51).AutoFit();

                            //Get "ki_hieu_hoa_don" and "so_hoa_don" from "tra_tien", match with "cong_no"
                            rs = db.OpenRecordset("tra_tien");
                            int maxTraTien = rs.RecordCount;
                            String paidID = String.Empty;
                            short month = 0;

                            //Liệt kê những khoản khách hàng trả thừa
                            ExcelWorksheet notMatchPaidWorksheet = package.Workbook.Worksheets.Add("Khach hang tra thua");
                            notMatchPaidWorksheet.Cells["A:E"].Style.Font.Name = "Times New Roman";
                            notMatchPaidWorksheet.Cells["A:E"].Style.Font.Size = 11;
                            notMatchPaidWorksheet.Cells["A:E"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            notMatchPaidWorksheet.View.FreezePanes(2, 1);

                            notMatchPaidWorksheet.Column(1).Width = GetTrueColumnWidth(8);
                            notMatchPaidWorksheet.Column(2).Width = GetTrueColumnWidth(11);
                            notMatchPaidWorksheet.Column(3).Width = GetTrueColumnWidth(12);
                            notMatchPaidWorksheet.Column(4).Width = GetTrueColumnWidth(15);
                            notMatchPaidWorksheet.Column(5).Width = GetTrueColumnWidth(18);

                            notMatchPaidWorksheet.Cells["E:E"].Style.Numberformat.Format = "_(* #,##0_);_(* (#,##0);_(* \" - \"_);_(@_)";

                            notMatchPaidWorksheet.Cells["A:D"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            notMatchPaidWorksheet.Cells["E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            notMatchPaidWorksheet.Cells["A1:E1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            notMatchPaidWorksheet.Cells["A1:E1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(247, 150, 70));
                            notMatchPaidWorksheet.Cells["A1:E1"].Style.Font.Bold = true;
                            notMatchPaidWorksheet.Cells["A1:E1"].Style.WrapText = true;

                            notMatchPaidWorksheet.Cells["A1"].Value = "STT";
                            notMatchPaidWorksheet.Cells["B1"].Value = "KÍ HIỆU HÓA ĐƠN";
                            notMatchPaidWorksheet.Cells["C1"].Value = "SỐ HÓA ĐƠN";
                            notMatchPaidWorksheet.Cells["D1"].Value = "THÁNG THANH TOÁN";
                            notMatchPaidWorksheet.Cells["E1"].Value = "TỔNG SỐ TIỀN THANH TOÁN";

                            int notMatchPaidCurrentRow = 1;

                            for (i = 1; i <= maxTraTien; i++)
                            {
                                paidID = String.Format("{0};{1}", rs.Fields["ki_hieu_hoa_don"].Value, rs.Fields["so_hoa_don"].Value);

                                if (invoices.ContainsKey(paidID))
                                {
                                    month = rs.Fields["thang_thanh_toan"].Value;
                                    switch (month)
                                    {
                                        case 1:
                                            worksheet.Cells["X" + invoices[paidID]].Value = rs.Fields["tong_thanh_toan"].Value;
                                            break;
                                        case 2:
                                            worksheet.Cells["Y" + invoices[paidID]].Value = rs.Fields["tong_thanh_toan"].Value;
                                            break;
                                        case 3:
                                            worksheet.Cells["Z" + invoices[paidID]].Value = rs.Fields["tong_thanh_toan"].Value;
                                            break;
                                        case 4:
                                            worksheet.Cells["AA" + invoices[paidID]].Value = rs.Fields["tong_thanh_toan"].Value;
                                            break;
                                        case 5:
                                            worksheet.Cells["AB" + invoices[paidID]].Value = rs.Fields["tong_thanh_toan"].Value;
                                            break;
                                        case 6:
                                            worksheet.Cells["AC" + invoices[paidID]].Value = rs.Fields["tong_thanh_toan"].Value;
                                            break;
                                        case 7:
                                            worksheet.Cells["AD" + invoices[paidID]].Value = rs.Fields["tong_thanh_toan"].Value;
                                            break;
                                        case 8:
                                            worksheet.Cells["AE" + invoices[paidID]].Value = rs.Fields["tong_thanh_toan"].Value;
                                            break;
                                        case 9:
                                            worksheet.Cells["AF" + invoices[paidID]].Value = rs.Fields["tong_thanh_toan"].Value;
                                            break;
                                        case 10:
                                            worksheet.Cells["AG" + invoices[paidID]].Value = rs.Fields["tong_thanh_toan"].Value;
                                            break;
                                        case 11:
                                            worksheet.Cells["AH" + invoices[paidID]].Value = rs.Fields["tong_thanh_toan"].Value;
                                            break;
                                        case 12:
                                            worksheet.Cells["AI" + invoices[paidID]].Value = rs.Fields["tong_thanh_toan"].Value;
                                            break;
                                    }
                                }
                                else
                                {
                                    notMatchPaidCurrentRow++;
                                    notMatchPaidWorksheet.Cells["A" + notMatchPaidCurrentRow].Value = notMatchPaidCurrentRow - 1;
                                    notMatchPaidWorksheet.Cells["B" + notMatchPaidCurrentRow].Value = rs.Fields["ki_hieu_hoa_don"].Value;
                                    notMatchPaidWorksheet.Cells["C" + notMatchPaidCurrentRow].Value = rs.Fields["so_hoa_don"].Value;
                                    notMatchPaidWorksheet.Cells["D" + notMatchPaidCurrentRow].Value = rs.Fields["thang_thanh_toan"].Value;
                                    notMatchPaidWorksheet.Cells["E" + notMatchPaidCurrentRow].Value = rs.Fields["tong_thanh_toan"].Value;
                                }

                                rs.MoveNext();
                            }

                            //Add total cells
                            worksheet.Row(rowTong).Height = 40;
                            worksheet.Row(rowTong).Style.Font.Size = 13;
                            worksheet.Cells["A" + rowTong + ":AT" + rowTong].Style.Font.Bold = true;

                            worksheet.Cells["D" + rowTong].Value = "Tổng cộng";

                            String fDuDauKy = String.Format("IF($F$1=\"Công ty\",SUM(J{0}:J{1}),SUMIF($B${0}:$B${1},$F$1,J{0}:J{1}))", ROW_BEFORE_START_EXCEL + 1, maxRowExcel);
                            worksheet.Cells["J" + rowTong].Formula = fDuDauKy;

                            //Copy formula
                            for (i = 10; i <= 45; i++)
                                worksheet.Cells["J" + rowTong].Copy(worksheet.Cells[rowTong, i]);

                            //Add border
                            worksheet.Cells["A" + ROW_BEFORE_START_EXCEL + ":AY" + rowTong].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A" + ROW_BEFORE_START_EXCEL + ":AY" + rowTong].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A" + ROW_BEFORE_START_EXCEL + ":AY" + rowTong].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheet.Cells["A" + ROW_BEFORE_START_EXCEL + ":AY" + rowTong].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                            //Content at end of report
                            int rowXacNhan = rowTong + 3;
                            worksheet.Row(rowXacNhan).Style.Numberformat.Format = "General";
                            worksheet.Cells["D" + rowXacNhan].Formula = "D3";

                            String fPhongXacNhan = String.Format("\"Xác nhận đối chiếu đến hết ngày \" & DAY(E1) & \" tháng \" & MONTH(E1) & \" năm \" & YEAR(E1)");
                            int rowDoiChieu = rowXacNhan + 1;
                            worksheet.Cells["D" + rowDoiChieu].Formula = fPhongXacNhan;

                            String fNgayThangNam = String.Format("\"Hải Phòng, ngày \" & DAY(E1) & \" tháng \" & MONTH(E1) & \" năm \" & YEAR(E1)");
                            worksheet.Cells["AQ" + (rowXacNhan - 1)].Formula = fNgayThangNam;

                            worksheet.Cells["AJ" + rowXacNhan].Value = "PHÒNG KẾ TOÁN";
                            worksheet.Cells["AQ" + rowXacNhan].Value = "LÃNH ĐẠO CÔNG TY";
                            worksheet.Cells["D" + rowXacNhan + ":AS" + rowXacNhan].Style.Font.Bold = true;

                            //Add CustomFilter VBA Project
                            package.Workbook.CreateVBAProject();
                            var sb = new StringBuilder();

                            sb.AppendLine("Private Sub Worksheet_SelectionChange(ByVal Target As Range)");
                            sb.AppendLine(" If Range(\"B1\") = \"ON\" Then");
                            sb.AppendLine("     If Range(\"F1\") = \"Công ty\" Then");
                            sb.AppendLine("         If Sheets(2).AutoFilterMode Then");
                            sb.AppendLine("             Sheets(2).AutoFilter.ShowAllData");
                            sb.AppendLine("         Else");
                            sb.AppendLine("             Range(\"A9:AY9\").AutoFilter");
                            sb.AppendLine("         End If");
                            sb.AppendLine("     Else");
                            sb.AppendLine("         Range(\"A9: AY9\").AutoFilter Field:=2, Criteria1:=Range(\"F1\").Value, Operator:=xlOr, Criteria2:=\"=\"");
                            sb.AppendLine("         Range(\"A9: AY9\").AutoFilter Field:=37, Criteria1:=\"<>0\"");
                            sb.AppendLine("     End If");
                            sb.AppendLine(" End If");
                            sb.AppendLine("End Sub");

                            worksheet.CodeModule.Code = sb.ToString();

                            //Add ON/OFF CustomFilter
                            worksheet.Cells["B1"].Style.Font.Bold = true;
                            worksheet.Cells["B1"].Style.Font.Size = 14;
                            worksheet.Cells["B1"].Style.Border.BorderAround(ExcelBorderStyle.Double);
                            worksheet.Cells["B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells["B1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(192, 80, 77));
                            worksheet.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            
                            var CustomFilterMode = worksheet.Cells["B1"].DataValidation.AddListDataValidation();
                            CustomFilterMode.Formula.Values.Add("ON");
                            CustomFilterMode.Formula.Values.Add("OFF");
                            worksheet.Cells["B1"].Value = "OFF";

                            //Conditional formatting Cell B1
                            ExcelAddress excelAddress = new ExcelAddress("B1");
                            String statement = "B1 = \"ON\"";
                            var condition = worksheet.ConditionalFormatting.AddExpression(excelAddress);
                            condition.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            condition.Style.Fill.BackgroundColor.Color = Color.FromArgb(155, 187, 89);
                            condition.Formula = statement;

                            //Filter, Scale, Freeze view
                            worksheet.Cells["A" + ROW_BEFORE_START_EXCEL + ":AY" + ROW_BEFORE_START_EXCEL].AutoFilter = true;
                            worksheet.View.FreezePanes(ROW_BEFORE_START_EXCEL + 1, 8);
                            worksheet.View.ZoomScale = 85;

                            //Sheet "Tong hop"
                            ExcelWorksheet worksheetTongHop = package.Workbook.Worksheets.Add("Tong hop");

                            worksheetTongHop.Column(1).Width = GetTrueColumnWidth(32);
                            for (i = 2; i <= 33; i++)
                            {
                                worksheetTongHop.Column(i).Width = GetTrueColumnWidth(25);
                            }

                            worksheetTongHop.Cells["A:AG"].Style.Font.Name = "Times New Roman";
                            worksheetTongHop.Cells["A:AG"].Style.Font.Size = 11;
                            worksheetTongHop.Cells["A:AG"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                            worksheetTongHop.Cells["A1:AG4"].Style.Font.Bold = true;
                            worksheetTongHop.Cells["A1:AG4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                            worksheetTongHop.Row(1).Height = 60;
                            worksheetTongHop.Row(2).Height = 60;
                            worksheetTongHop.Row(3).Height = 40;

                            worksheetTongHop.Cells["A1:AG1"].Merge = true;
                            String fBaoCaoTongHop = String.Format("\"BÁO CÁO NỢ PHẢI THU KHÁCH HÀNG ĐẾN NGÀY \" & DAY('Doi chieu cong no'!E1) & \" THÁNG \" & MONTH('Doi chieu cong no'!E1) & \" NĂM \" & YEAR('Doi chieu cong no'!E1)");
                            worksheetTongHop.Cells["A1:AG1"].Formula = fBaoCaoTongHop;

                            worksheetTongHop.Cells["A2:A4"].Merge = true;
                            worksheetTongHop.Cells["A2:A4"].Value = "TÊN PHÒNG";

                            worksheetTongHop.Cells["B2:B3"].Merge = true;
                            worksheetTongHop.Cells["B2:B3"].Value = "NỢ ĐẦU NĂM";

                            worksheetTongHop.Cells["C2:O2"].Merge = true;
                            worksheetTongHop.Cells["C2:O2"].Value = "NỢ PHÁT SINH";

                            worksheetTongHop.Cells["P2:AB2"].Merge = true;
                            worksheetTongHop.Cells["P2:AB2"].Value = "THU NỢ";

                            worksheetTongHop.Cells["AC2:AF2"].Merge = true;
                            worksheetTongHop.Cells["AC2:AF2"].Value = "CHI TIẾT TUỔI NỢ";

                            worksheetTongHop.Cells["AG2:AG3"].Merge = true;
                            worksheetTongHop.Cells["AG2:AG3"].Value = "TỔNG NỢ CUỐI KỲ";
                        
                            int monthTongHop = 1;
                            for (i = 3; i <= 14; i++)
                            {
                                worksheetTongHop.Cells[3, i].Value = "THÁNG " + monthTongHop.ToString();
                                monthTongHop++;
                            }

                            monthTongHop = 1;
                            for (i = 16; i <= 27; i++)
                            {
                                worksheetTongHop.Cells[3, i].Value = "THÁNG " + monthTongHop.ToString();
                                monthTongHop++;
                            }

                            worksheetTongHop.Cells["O3"].Value = "TỔNG CỘNG";
                            worksheetTongHop.Cells["AB3"].Value = "TỔNG CỘNG";
                            worksheetTongHop.Cells["AC3"].Value = "NỢ ĐẾN HẠN TT";
                            worksheetTongHop.Cells["AD3"].Value = "NỢ QHTT DƯỚI 1 NĂM";
                            worksheetTongHop.Cells["AE3"].Value = "NỢ QHTT TỪ 1-2 NĂM";
                            worksheetTongHop.Cells["AF3"].Value = "NỢ QHTT TRÊN 2 NĂM";
                            worksheetTongHop.Cells[4, 2, 4, 33].Value = "VND";

                            int rowOfTongHopWorksheet = 5;

                            String tenPhongVietTat;
                            String tenPhongDayDu;
                            String fSumIf;
                            String fSumNo;

                            foreach (KeyValuePair<String, String> department in departments)
                            {
                                worksheetTongHop.Row(rowOfTongHopWorksheet).Height = 45;
                                tenPhongVietTat = department.Key;
                                tenPhongDayDu = department.Value;

                                worksheetTongHop.Cells["A" + rowOfTongHopWorksheet.ToString()].Value = tenPhongDayDu;
                                worksheetTongHop.Cells["A" + rowOfTongHopWorksheet].Style.Font.Bold = true;

                                fSumIf = String.Format("SUMIF('Doi chieu cong no'!$B${0}:$B${1}, \"{2}\",'Doi chieu cong no'!J{0}:J{1})", ROW_BEFORE_START_EXCEL + 1, maxRowExcel, tenPhongVietTat);
                                worksheetTongHop.Cells["B" + rowOfTongHopWorksheet.ToString()].Formula = fSumIf;
                                for (int j = 3; j <= 28; j++)
                                    worksheetTongHop.Cells["B" + rowOfTongHopWorksheet.ToString()].Copy(worksheetTongHop.Cells[rowOfTongHopWorksheet, j]);

                                fSumIf = String.Format("SUMIF('Doi chieu cong no'!$B${0}:$B${1}, \"{2}\",'Doi chieu cong no'!AL{0}:AL{1})", ROW_BEFORE_START_EXCEL + 1, maxRowExcel, tenPhongVietTat);
                                worksheetTongHop.Cells["AC" + rowOfTongHopWorksheet.ToString()].Formula = fSumIf;

                                //Duoi 1 nam
                                fSumIf = String.Format("SUMIF('Doi chieu cong no'!$B${0}:$B${1}, \"{2}\",'Doi chieu cong no'!AM{0}:AM{1}) " +
                                                    "+ SUMIF('Doi chieu cong no'!$B${0}:$B${1}, \"{2}\",'Doi chieu cong no'!AN{0}:AN{1}) " +
                                                    "+ SUMIF('Doi chieu cong no'!$B${0}:$B${1}, \"{2}\",'Doi chieu cong no'!AO{0}:AO{1}) " +
                                                    "+ SUMIF('Doi chieu cong no'!$B${0}:$B${1}, \"{2}\",'Doi chieu cong no'!AP{0}:AP{1})"
                                                        , ROW_BEFORE_START_EXCEL + 1, maxRowExcel, tenPhongVietTat);
                                worksheetTongHop.Cells["AD" + rowOfTongHopWorksheet.ToString()].Formula = fSumIf;

                                //Tu 1-2 nam
                                fSumIf = String.Format("SUMIF('Doi chieu cong no'!$B${0}:$B${1}, \"{2}\",'Doi chieu cong no'!AQ{0}:AQ{1})", ROW_BEFORE_START_EXCEL + 1, maxRowExcel, tenPhongVietTat);
                                worksheetTongHop.Cells["AE" + rowOfTongHopWorksheet.ToString()].Formula = fSumIf;

                                //Tren 2 nam
                                fSumIf = String.Format("SUMIF('Doi chieu cong no'!$B${0}:$B${1}, \"{2}\",'Doi chieu cong no'!AR{0}:AR{1}) " +
                                                    "+ SUMIF('Doi chieu cong no'!$B${0}:$B${1}, \"{2}\",'Doi chieu cong no'!AS{0}:AS{1}) "
                                                        , ROW_BEFORE_START_EXCEL + 1, maxRowExcel, tenPhongVietTat);
                                worksheetTongHop.Cells["AF" + rowOfTongHopWorksheet.ToString()].Formula = fSumIf;

                                //Tong no cuoi ky
                                fSumNo = String.Format("B{0} + O{0} - AB{0}", rowOfTongHopWorksheet);
                                worksheetTongHop.Cells["AG" + rowOfTongHopWorksheet.ToString()].Formula = fSumNo;

                                rowOfTongHopWorksheet++;
                            }

                            int rowTotalOfTongHop = rowOfTongHopWorksheet;
                            worksheetTongHop.Row(rowTotalOfTongHop).Height = 55;
                            worksheetTongHop.Cells[rowTotalOfTongHop, 1, rowTotalOfTongHop, 33].Style.Font.Bold = true;
                            worksheetTongHop.Cells[rowTotalOfTongHop, 1, rowTotalOfTongHop, 33].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheetTongHop.Cells[rowTotalOfTongHop, 1, rowTotalOfTongHop, 33].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 204, 153));

                            worksheetTongHop.Cells["A" + rowTotalOfTongHop].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheetTongHop.Cells["A" + rowTotalOfTongHop].Value = "CỘNG";

                            String fSumTongHop;
                            fSumTongHop = String.Format("SUM(B5:B{0})", rowOfTongHopWorksheet - 1);
                            worksheetTongHop.Cells["B" + rowTotalOfTongHop].Formula = fSumTongHop;

                            for (i = 3; i <= 33; i++)
                                worksheetTongHop.Cells["B" + rowTotalOfTongHop].Copy(worksheetTongHop.Cells[rowTotalOfTongHop, i]);

                            worksheetTongHop.Cells[5, 2, rowTotalOfTongHop, 33].Style.Numberformat.Format = "_(* #,##0_);_(* (#,##0);_(* \" - \"_);_(@_)";
                            worksheetTongHop.Cells[2, 1, rowTotalOfTongHop, 33].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            worksheetTongHop.Cells[2, 1, rowTotalOfTongHop, 33].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            worksheetTongHop.Cells[2, 1, rowTotalOfTongHop, 33].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            worksheetTongHop.Cells[2, 1, rowTotalOfTongHop, 33].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            worksheetTongHop.Cells[2, 1, rowTotalOfTongHop, 33].Style.Border.BorderAround(ExcelBorderStyle.Double);
                            

                            //Save and message Done
                            package.SaveAs(newFile);

                            MessageBox.Show("Đã lập đối chiếu công nợ", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            if (loaiTk.Text == "Tổng")
                            {
                                //Get outstanding from "cong_no", write to json file
                                worksheet.Cells["AK" + ROW_BEFORE_START_EXCEL + ":AK" + maxRowExcel].Calculate();
                                rs = db.OpenRecordset("cong_no");
                                if (!rs.BOF)
                                    rs.MoveFirst();

                                double debt = 0;

                                //Write to Json
                                StringBuilder sbOutstanding = new StringBuilder();
                                StringWriter swOutstanding = new StringWriter(sbOutstanding);

                                using (JsonWriter writerOutstanding = new JsonTextWriter(swOutstanding))
                                {
                                    writerOutstanding.Formatting = Formatting.Indented;
                                    writerOutstanding.WriteStartArray();

                                    for (i = ROW_BEFORE_START_EXCEL + 1; i <= maxRowExcel; i++)
                                    {
                                        debt = Convert.ToDouble(worksheet.Cells["AK" + i].Value);

                                        if (debt > 0)
                                        {
                                            writerOutstanding.WriteStartObject();

                                            writerOutstanding.WritePropertyName("KiHieuHoaDon");
                                            writerOutstanding.WriteValue(rs.Fields["ki_hieu_hoa_don"].Value);

                                            writerOutstanding.WritePropertyName("SoHoaDon");
                                            writerOutstanding.WriteValue(rs.Fields["so_hoa_don"].Value);

                                            writerOutstanding.WritePropertyName("MST");
                                            writerOutstanding.WriteValue(rs.Fields["m_mst"].Value);

                                            writerOutstanding.WritePropertyName("KhachHang");
                                            writerOutstanding.WriteValue(rs.Fields["m_cong_ty"].Value);

                                            writerOutstanding.WritePropertyName("HanThanhToan");
                                            writerOutstanding.WriteValue(rs.Fields["m_han_thanh_toan"].Value);

                                            writerOutstanding.WritePropertyName("SoTienPhatSinh");
                                            writerOutstanding.WriteValue(worksheet.Cells["AK" + i].Value);

                                            writerOutstanding.WritePropertyName("NgayChungTu");
                                            writerOutstanding.WriteValue(rs.Fields["m_ngay_ct"].Value);

                                            writerOutstanding.WritePropertyName("NgayHoaDon");
                                            writerOutstanding.WriteValue(rs.Fields["m_ngay_hoa_don"].Value);

                                            writerOutstanding.WritePropertyName("MaPhong");
                                            writerOutstanding.WriteValue(rs.Fields["m_ma_phong"].Value);

                                            writerOutstanding.WritePropertyName("MaNghiepVu");
                                            writerOutstanding.WriteValue(rs.Fields["m_ma_nv"].Value);

                                            writerOutstanding.WritePropertyName("User");
                                            writerOutstanding.WriteValue(rs.Fields["m_user_nhap"].Value);

                                            writerOutstanding.WritePropertyName("KenhKT");
                                            writerOutstanding.WriteValue(rs.Fields["m_kenh_kt"].Value);

                                            writerOutstanding.WritePropertyName("SoTk");
                                            writerOutstanding.WriteValue(rs.Fields["m_so_tai_khoan"].Value);

                                            writerOutstanding.WritePropertyName("SoBk");
                                            writerOutstanding.WriteValue(rs.Fields["m_so_tham_chieu"].Value);

                                            writerOutstanding.WritePropertyName("LoaiTien");
                                            writerOutstanding.WriteValue(rs.Fields["m_loai_tien"].Value);

                                            writerOutstanding.WritePropertyName("NguyenTe");
                                            writerOutstanding.WriteValue(rs.Fields["m_tong_nguyen_te"].Value);

                                            writerOutstanding.WritePropertyName("MaCanBo");
                                            writerOutstanding.WriteValue(rs.Fields["m_id"].Value);

                                            writerOutstanding.WriteEndObject();
                                        }
                                        rs.MoveNext();
                                    }
                                    writerOutstanding.WriteEndArray();
                                }

                                String ToJsonOutstanding = swOutstanding.ToString();

                                DirectoryInfo directoryInfo = new DirectoryInfo(Environment.CurrentDirectory + @"\json");
                                if (!directoryInfo.Exists)
                                    Directory.CreateDirectory(Environment.CurrentDirectory + @"\json");
                                int NextYear = Convert.ToInt32(Program.DbYear) + 1;
                                String outstandingToObject = Environment.CurrentDirectory + @"\json\outstanding - " + NextYear.ToString() + ".json";
                                if (File.Exists(outstandingToObject))
                                    File.Delete(outstandingToObject);

                                using (StreamWriter outstandingToJson = new StreamWriter(outstandingToObject))
                                {
                                    outstandingToJson.WriteLine(ToJsonOutstanding);
                                }
                            }
                        }
                        rs.Close();
                        db.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể lập báo cáo.\n" + ex.Message.ToString(), "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Search_Click(object sender, EventArgs e)
        {
            searchList.Enabled = true;
            searchList.Rows.Clear();
            CurrentInfoRefresh(RefreshOption.All);
            String searchWhat = tbSearch.Text.Trim();

            String db_name = Program.DbYear + ".mdb";
            String db_path = Environment.CurrentDirectory + @"\Database\";
            String db_file = db_path + db_name;

            DBEngine dBEngine = new DBEngine();
            Database db;
            Recordset rs = null;

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
                                searchValue = Convert.ToString(rs.Fields["so_hoa_don"].Value);
                                matchSearchCondition = searchValue.Contains(searchWhat);
                                if (matchSearchCondition)
                                {
                                    ten_don_vi = rs.Fields["mst"].Value;
                                    mst = rs.Fields["mst"].Value;
                                    ma_hoa_don = rs.Fields["ki_hieu_hoa_don"].Value;
                                    so_hoa_don = rs.Fields["so_hoa_don"].Value;
                                }
                                if (rs.Name == "invoice")
                                {
                                    so_tien = Convert.ToDouble(rs.Fields["so_tien_phat_sinh"].Value);
                                }
                                else
                                {
                                    so_tien = Convert.ToDouble(rs.Fields["so_tien_thanh_toan"].Value);
                                }
                                break;
                            case "Số tiền":
                                if (rs.Name == "invoice")
                                    searchDouble = Convert.ToDouble(rs.Fields["so_tien_phat_sinh"].Value);
                                else
                                    searchDouble = Convert.ToDouble(rs.Fields["so_tien_thanh_toan"].Value);
                                
                                matchSearchCondition = searchDouble == Convert.ToDouble(searchWhat);
                                if (matchSearchCondition)
                                {
                                    mst = rs.Fields["mst"].Value;
                                    ma_hoa_don = rs.Fields["ki_hieu_hoa_don"].Value;
                                    so_hoa_don = rs.Fields["so_hoa_don"].Value;
                                    if (rs.Name == "invoice")
                                        so_tien = Convert.ToDouble(rs.Fields["so_tien_phat_sinh"].Value);
                                    else
                                        so_tien = Convert.ToDouble(rs.Fields["so_tien_thanh_toan"].Value);
                                }
                                break;
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

            String db_name = Program.DbYear + ".mdb";
            String db_path = Environment.CurrentDirectory + @"\Database\";
            String db_file = db_path + db_name;

            DBEngine dBEngine = new DBEngine();
            Database db;
            Recordset rs = null;

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
                            rs = db.OpenRecordset("invoice");
                        else if (field == "Thu nợ")
                            rs = db.OpenRecordset("paid");

                        if (!rs.BOF)
                            rs.MoveFirst();

                        rs.Move(recordNumber);

                        currentMST.Text = rs.Fields["mst"].Value;
                        Recordset rsKhachHang = db.OpenRecordset("SELECT cong_ty FROM customers WHERE mst = '" + currentMST.Text + "'");
                        if (rsKhachHang.RecordCount > 0)
                            currentKhachHang.Text = rsKhachHang.Fields["cong_ty"].Value;

                        currentMaHoaDon.Text = rs.Fields["ki_hieu_hoa_don"].Value;
                        currentSoHoaDon.Text = rs.Fields["so_hoa_don"].Value;
                        currentHanTra.Text = String.Format("{0:dd/MM/yyyy}", rs.Fields["han_thanh_toan"].Value);
                        
                        currentNgayHoaDon.Text = String.Format("{0:dd/MM/yyyy}", rs.Fields["ngay_hoa_don"].Value);

                        if (field == "Thu nợ")
                        {
                            currentNgayTra.Text = String.Format("{0:dd/MM/yyyy}", rs.Fields["ngay_ct"].Value);
                            currentSoTienTra.Text = String.Format("{0:n0}", rs.Fields["so_tien_thanh_toan"].Value);
                            currentNgayTra.Text = String.Format("{0:dd/MM/yyyy}", rs.Fields["ngay_ct"].Value);
                        }
                        else if (field == "Phát sinh")
                        {
                            currentSoTienNo.Text = String.Format("{0:n0}", rs.Fields["so_tien_phat_sinh"].Value);
                            currentNgayPhatSinh.Text = String.Format("{0:dd/MM/yyyy}", rs.Fields["ngay_ct"].Value);
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

                //Create "Mau lay du lieu Sunweb" form
                if (!File.Exists(newFile.ToString()))
                {
                    using (var package = new ExcelPackage(newFile))
                    {
                        package.Workbook.Properties.Title = "Mau lay du lieu Sunweb";
                        package.Workbook.Properties.Company = "Bảo Việt Hải Phòng";

                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Du lieu Sunweb");
                        worksheet.Cells["A:T"].Style.Font.Name = "Calibri";
                        worksheet.Cells["A1:T1"].Style.Font.Bold = true;

                        worksheet.Column(1).Width = GetTrueColumnWidth(6);
                        worksheet.Column(2).Width = GetTrueColumnWidth(9);
                        worksheet.Column(3).Width = GetTrueColumnWidth(9);
                        worksheet.Column(4).Width = GetTrueColumnWidth(12);
                        worksheet.Column(5).Width = GetTrueColumnWidth(5);
                        worksheet.Column(6).Width = GetTrueColumnWidth(10);
                        worksheet.Column(7).Width = GetTrueColumnWidth(35);
                        worksheet.Column(8).Width = GetTrueColumnWidth(35);
                        worksheet.Column(9).Width = GetTrueColumnWidth(30);
                        worksheet.Column(10).Width = GetTrueColumnWidth(6);
                        worksheet.Column(11).Width = GetTrueColumnWidth(12);
                        worksheet.Column(12).Width = GetTrueColumnWidth(11);
                        worksheet.Column(13).Width = GetTrueColumnWidth(8);
                        worksheet.Column(14).Width = GetTrueColumnWidth(8.5);
                        worksheet.Column(15).Width = GetTrueColumnWidth(13);
                        worksheet.Column(16).Width = GetTrueColumnWidth(13);
                        worksheet.Column(17).Width = GetTrueColumnWidth(9);
                        worksheet.Column(18).Width = GetTrueColumnWidth(15);
                        worksheet.Column(19).Width = GetTrueColumnWidth(15.20);

                        worksheet.Cells["A:T"].Style.Font.Size = 8;
                        worksheet.Cells["A:T"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["A:T"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["A:T"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["A:T"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        worksheet.Cells["A1"].Value = "LOAICT";
                        worksheet.Cells["B1"].Value = "NGAYCT";
                        worksheet.Cells["C1"].Value = "SOTAIKHOAN";
                        worksheet.Cells["D1"].Value = "SOTIENQUYDOI";
                        worksheet.Cells["E1"].Value = "NOCO";
                        worksheet.Cells["F1"].Value = "SOTHAMCHIEU";
                        worksheet.Cells["G1"].Value = "GNRL_DESCR_01";
                        worksheet.Cells["H1"].Value = "GNRL_DESCR_02";
                        worksheet.Cells["I1"].Value = "NGAYDAOHAN";
                        worksheet.Cells["J1"].Value = "LOAITIEN";
                        worksheet.Cells["K1"].Value = "SOTIEN";
                        worksheet.Cells["L1"].Value = "T2";
                        worksheet.Cells["M1"].Value = "T3";
                        worksheet.Cells["N1"].Value = "T5";
                        worksheet.Cells["O1"].Value = "T8";
                        worksheet.Cells["P1"].Value = "MASOTHUE";
                        worksheet.Cells["Q1"].Value = "KYHIEUHOADON";
                        worksheet.Cells["R1"].Value = "SOHOADON";
                        worksheet.Cells["S1"].Value = "NGAYHOADONGOC";
                        worksheet.Cells["T1"].Value = "USERNHAP";
                        

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

            afterMST.Text = currentMST.Text;
            afterKhachHang.Text = currentKhachHang.Text;
            afterMaHoaDon.Text = currentMaHoaDon.Text;
            afterSoHoaDon.Text = currentSoHoaDon.Text;
            afterHanTra.Text = currentHanTra.Text;
            afterSoTienNo.Text = currentSoTienNo.Text;
            afterNgayPhatSinh.Text = currentNgayPhatSinh.Text;
            afterNgayHoaDon.Text = currentNgayHoaDon.Text;
            afterNgayTra.Text = currentNgayTra.Text;
            afterSoTienTra.Text = currentSoTienTra.Text;
        }

        private void Delete_Click(object sender, EventArgs e)
        {
            String db_name = Program.DbYear + ".mdb";
            String db_path = Environment.CurrentDirectory + @"\Database\";
            String db_file = db_path + db_name;

            DBEngine dBEngine = new DBEngine();
            Database db;
            Recordset rs = null;
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
                String db_name = Program.DbYear + ".mdb";
                String db_path = Environment.CurrentDirectory + @"\Database\";
                String db_file = db_path + db_name;

                DBEngine dBEngine = new DBEngine();
                Database db;
                Recordset rs = null;
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
                                rs = db.OpenRecordset("invoice");
                            else if (field == "Thu nợ")
                                rs = db.OpenRecordset("paid");

                            if (!rs.BOF)
                                rs.MoveFirst();

                            rs.Move(recordNumber);
                            rs.Edit();
                            rs.Fields["mst"].Value = afterMST.Text;
                            rs.Fields["ki_hieu_hoa_don"].Value = afterMaHoaDon.Text;
                            rs.Fields["so_hoa_don"].Value = afterSoHoaDon.Text;
                            rs.Fields["han_thanh_toan"].Value = afterHanTra.Text;
                            rs.Fields["ngay_hoa_don"].Value = afterNgayHoaDon.Text;

                            if (field == "Phát sinh")
                            {
                                rs.Fields["so_tien_phat_sinh"].Value = afterSoTienNo.Text;
                                rs.Fields["ngay_ct"].Value = afterNgayPhatSinh.Text;
                            }
                            else if (field == "Thu nợ")
                            {
                                rs.Fields["so_tien_thanh_toan"].Value = afterSoTienTra.Text;
                                rs.Fields["ngay_ct"].Value = afterNgayTra.Text;
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
            afterMST.Clear();
            afterKhachHang.Clear();
            afterMaHoaDon.Clear();
            afterSoHoaDon.Clear();
            afterHanTra.ResetText();
            afterSoTienTra.Clear();
            afterNgayPhatSinh.ResetText();
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
                        afterMST.Enabled = true;
                        afterMST.BackColor = SystemColors.Window;

                        afterMaHoaDon.Enabled = true;
                        afterMaHoaDon.BackColor = SystemColors.Window;

                        afterSoHoaDon.Enabled = true;
                        afterSoHoaDon.BackColor = SystemColors.Window;

                        afterHanTra.Enabled = true;
                        afterHanTra.BackColor = SystemColors.Window;

                        afterNgayHoaDon.Enabled = true;
                        afterNgayHoaDon.BackColor = SystemColors.Window;

                        if (field == "Phát sinh")
                        {
                            afterNgayPhatSinh.Enabled = true;
                            afterNgayPhatSinh.BackColor = SystemColors.Window;

                            afterSoTienNo.Enabled = true;
                            afterSoTienNo.BackColor = SystemColors.Window;
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

        private void NextYear_Click(object sender, EventArgs e)
        {
            try
            {
                int NextYear = Convert.ToInt32(Program.DbYear) + 1;
                String DBFileName = NextYear.ToString() + ".mdb";

                DialogResult dialogResult = MessageBox.Show("Bạn chắc chắn muốn tổng kết số liệu năm " + Program.DbYear +
                    " và lập dữ liệu mới của năm " + NextYear.ToString() + " ?", "Tổng kết năm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dialogResult == DialogResult.Yes)
                {
                    DBEngine dBEngine = new DBEngine();
                    Database db = null;
                    Recordset rsInvoice = null;
                    Recordset rsRevenue = null;
                    Recordset rsCustomers = null;

                    //Read Json
                    String outstandingReaderPath = Environment.CurrentDirectory + @"\json\outstanding - " + NextYear.ToString() + ".json";

                    if (File.Exists(outstandingReaderPath))
                    {
                        FileInfo fileInfo = new FileInfo(Environment.CurrentDirectory + @"\Database\" + DBFileName);
                        String db_file = fileInfo.ToString();

                        //Create new DB of next year
                        if (File.Exists(db_file))
                            File.Delete(db_file);

                        using (Stream s = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("CongNo.DB.mdb"))
                        {
                            using (FileStream ResourceFile = new FileStream(fileInfo.ToString(), FileMode.Create, FileAccess.Write))
                            {
                                s.CopyTo(ResourceFile);
                            }
                        }

                        db = dBEngine.OpenDatabase(db_file);
                        String queryName = "cong_no_draft";
                        String querySql = String.Format("SELECT invoice.ngay_ct, invoice.ngay_hoa_don, department.ma_phong," +
                            " department.ten_phong, customers.mst, customers.cong_ty, " +
                            "invoice.ki_hieu_hoa_don, invoice.so_hoa_don, invoice.han_thanh_toan, revenue.ma_nv, revenue.user_nhap, invoice.kenh_kt, " +
                            "revenue.so_tai_khoan, revenue.so_tham_chieu, invoice.loai_tien, invoice.tong_nguyen_te, " +
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
                            " ON department.ma_phong = revenue.ma_phong ORDER BY invoice.ki_hieu_hoa_don, invoice.so_hoa_don;", NextYear.ToString());

                        QueryDef cong_no_draft = new QueryDef();
                        cong_no_draft.Name = queryName;
                        cong_no_draft.SQL = querySql;

                        db.QueryDefs.Append(cong_no_draft);

                        rsInvoice = db.OpenRecordset("invoice");
                        rsRevenue = db.OpenRecordset("revenue");
                        rsCustomers = db.OpenRecordset("customers");

                        //Export data from "outstanding.json" to "invoice" and "revenue" table
                        using (StreamReader rOutstanding = new StreamReader(outstandingReaderPath))
                        {
                            String jsonOutstandingRead = rOutstanding.ReadToEnd();
                            dynamic jObjectOutstandings = JsonConvert.DeserializeObject(jsonOutstandingRead);

                            foreach (var outstanding in jObjectOutstandings)
                            {
                                rsInvoice.AddNew();
                                rsInvoice.Fields["ki_hieu_hoa_don"].Value = outstanding.KiHieuHoaDon;
                                rsInvoice.Fields["so_hoa_don"].Value = outstanding.SoHoaDon;
                                rsInvoice.Fields["mst"].Value = outstanding.MST;
                                rsInvoice.Fields["han_thanh_toan"].Value = outstanding.HanThanhToan;
                                rsInvoice.Fields["so_tien_phat_sinh"].Value = outstanding.SoTienPhatSinh;
                                rsInvoice.Fields["ngay_ct"].Value = outstanding.NgayChungTu;
                                rsInvoice.Fields["ngay_hoa_don"].Value = outstanding.NgayHoaDon;
                                rsInvoice.Fields["kenh_kt"].Value = outstanding.KenhKT;
                                rsInvoice.Fields["loai_tien"].Value = outstanding.LoaiTien;
                                rsInvoice.Fields["tong_nguyen_te"].Value = outstanding.NguyenTe;
                                rsInvoice.Update();

                                rsRevenue.AddNew();
                                rsRevenue.Fields["ki_hieu_hoa_don"].Value = outstanding.KiHieuHoaDon;
                                rsRevenue.Fields["so_hoa_don"].Value = outstanding.SoHoaDon;
                                rsRevenue.Fields["ma_phong"].Value = outstanding.MaPhong;
                                rsRevenue.Fields["ma_nv"].Value = outstanding.MaNghiepVu;
                                rsRevenue.Fields["user_nhap"].Value = outstanding.User;
                                rsRevenue.Fields["so_tai_khoan"].Value = outstanding.SoTk;
                                rsRevenue.Fields["so_tham_chieu"].Value = outstanding.SoBk;
                                rsRevenue.Update();

                                try
                                {
                                    rsCustomers.AddNew();
                                    rsCustomers.Fields["mst"].Value = outstanding.MST;
                                    rsCustomers.Fields["cong_ty"].Value = outstanding.KhachHang;
                                    rsCustomers.Update();
                                }
                                catch { }
                            }
                        }
                        rsInvoice.Close();
                        rsRevenue.Close();
                        rsCustomers.Close();
                        db.Close();

                        MessageBox.Show("Đã lập dữ liệu năm mới", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Không có dữ liệu, hãy kết xuất lại dữ liệu", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể lập dữ liệu năm mới.\n" + ex.Message.ToString(), "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void compactButton_Click(object sender, EventArgs e)
        {
            String db_name = Program.DbYear + ".mdb";
            String db_path = Environment.CurrentDirectory + @"\Database\";
            String db_file = db_path + db_name;
            DBEngine dBEngine = new DBEngine();

            try
            {
                //Compact Database
                String compactDbTemp = db_path + "temp.mdb";
                String compactDbName = db_path + Program.DbYear + ".mdb";
                dBEngine.CompactDatabase(db_file, compactDbTemp, null);
                File.Delete(db_file);
                File.Move(compactDbTemp, compactDbName);

                MessageBox.Show("Đã nén xong dữ liệu");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không nén được dữ liệu.\n" + ex.Message.ToString(), "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }
    }
}