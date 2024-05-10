using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace QL_NCKH
{
    public partial class uc_giangvien : UserControl
    {
        MyClass MyClass = new MyClass();
        XoaGV xoa = new XoaGV();
        //MenuMain menu = new MenuMain();
        public uc_giangvien()
        {
            InitializeComponent();
        }
        public void LoadDL()
        {
            string query = "select * from GiangVien ";
            DataTable dt = MyClass.DocDL(query);
            dgv_gv.DataSource = dt;
            dgv_gv.Columns[0].HeaderText = "Mã giảng viên";
            dgv_gv.Columns[1].HeaderText = "Họ Tên";
            dgv_gv.Columns[2].HeaderText = "Ngày Sinh";
            dgv_gv.Columns[3].HeaderText = "Địa chỉ";
            dgv_gv.Columns[4].HeaderText = "Học vị";
            dgv_gv.Columns[5].HeaderText = "Học hàm";
            dgv_gv.Columns[6].HeaderText = "Khoa chủ quản";
            dgv_gv.Columns[7].HeaderText = "Email";
            dgv_gv.Columns[8].HeaderText = "Số điện thoại";
            dgv_gv.Columns[9].HeaderText = "Giới tính";
            dgv_gv.Columns[10].HeaderText = "Đơn vị công tác";
            dgv_gv.Columns[10].Width = 200;

        }
        private void UserControl1_Load(object sender, EventArgs e)
        {
            LoadDL();
            
        }

        private void btn_thoat_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }
        public bool kiemtra()
        {
            if (string.IsNullOrWhiteSpace(txt_ma.Text) || string.IsNullOrWhiteSpace(txt_hoten.Text)
                || string.IsNullOrWhiteSpace(txt_diachi.Text) || string.IsNullOrWhiteSpace(txt_hocvi.Text)
                || string.IsNullOrWhiteSpace(txt_hocham.Text) || string.IsNullOrWhiteSpace(dtp_ngaysinh.Text)
                || string.IsNullOrWhiteSpace(cb_khoa.Text) || string.IsNullOrWhiteSpace(txt_email.Text)
                || string.IsNullOrWhiteSpace(txt_sdt.Text) || string.IsNullOrWhiteSpace(cb_gioitinh.Text) || string.IsNullOrWhiteSpace(txt_donvi.Text))
                return false;

                return true;
        }
        public bool kiemTraMa(string ma)
        {
            try
            {
                string sql = "select MaGV from GiangVien Where MaGV = '"+ma+"' ";
                DataTable tb = MyClass.DocDL(sql);
                if(tb.Rows.Count > 0)
                {
                    return false;
                }
               
            }
            catch
            {
                MessageBox.Show("Lỗi không kiểm tra được mã", "Thông báo");

            }
            return true;
        }

        public bool IsValidEmail(string email)
        {
            string pattern = @"^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$";
            Regex regex = new Regex(pattern);
            return regex.IsMatch(email);
        }
        public bool kiemtraSoDT(string sodt)
        {
            sodt = sodt.Replace(" ","");
            List<string> dienthoai = new List<string> { "03","08","09","07","05"};
            int count = 0;
            foreach (var item in sodt)
            {
                if(!char.IsDigit(item))
                {
                    count++;
                }
            }
            if(count == 0 && sodt.Count() < 10)
            {
                return false;
            }
            else
            {
                int count1 = 0;
                string dauso = sodt[0].ToString() + sodt[1].ToString();
                for (int i = 0; i < dienthoai.Count; i++)
                {
                    if (dauso == dienthoai[i])
                    {
                        count1++;
                    }
                   
                }

                if (count1 == 0)
                    return false;


                
            }
            return true;
                
            
        }
        
        public void ExcelExport()
        {
            try
            {
                Excel.Application oExcel = new Excel.Application();
                Excel.Workbook oBook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

                Excel.Worksheet oSheet = (Excel.Worksheet)oBook.Worksheets[1];

                Excel.Range head = oSheet.get_Range("A1", "K1");

                head.MergeCells = true;

                head.Value2 = "DANH SÁCH GIẢNG VIÊN";

                head.Font.Bold = true;

                head.Font.Name = "Times New Roman";

                head.Font.Size = "20";

                head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range cl1 = oSheet.get_Range("A3","A3");
                cl1.Value = "Mã giảng viên";

                Excel.Range cl2 = oSheet.get_Range("B3","B3");
                cl2.Value = "Họ Tên";
                Excel.Range cl3 = oSheet.get_Range("C3", "C3");
                cl3.Value = "Ngày sinh";

                Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                cl4.Value = "Địa chỉ";

                Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                cl5.Value = "Học vị";

                Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                cl6.Value = "Học hàm";
                Excel.Range cl7 = oSheet.get_Range("G3", "G3");
                cl7.Value = "Khoa chủ quản";

                Excel.Range cl8 = oSheet.get_Range("H3", "H3");
                cl8.Value = "Email";

                Excel.Range cl9 = oSheet.get_Range("I3", "I3");
                cl9.Value = "Số điện thoại";

                Excel.Range cl10 = oSheet.get_Range("J3", "J3");
                cl10.Value = "Giới tính";

                Excel.Range cl11 = oSheet.get_Range("K3", "K3");
                cl11.Value = "Đơn vị công tác";

                Excel.Range rowHead = oSheet.get_Range("A3", "K3");
                rowHead.Font.Bold = true;
                rowHead.Font.Size = 13;
                rowHead.Font.Name = "Times New Roman";
                rowHead.Borders.LineStyle = Excel.Constants.xlSolid;
                rowHead.Interior.ColorIndex = 6;
                rowHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int line = 4;
                for (int i = 0; i < dgv_gv.Rows.Count - 1; i++)
                {
                    Excel.Range line1 = oSheet.get_Range("A" + (line + i).ToString(),"A" + (line + i).ToString());
                    line1.Value = dgv_gv.Rows[i].Cells[0].Value.ToString();
                    line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line1.Borders.LineStyle = Excel.Constants.xlSolid;
                    line1.Font.Name = "Times New Roman";

                    Excel.Range line2= oSheet.get_Range("B" + (line + i).ToString(), "B" + (line + i).ToString());
                    line2.Value = dgv_gv.Rows[i].Cells[1].Value.ToString();
                    line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line2.Borders.LineStyle = Excel.Constants.xlSolid;
                    line2.Font.Name = "Times New Roman";

                    Excel.Range line3 = oSheet.get_Range("C" + (line + i).ToString(), "C" + (line + i).ToString());
                    line3.Value = dgv_gv.Rows[i].Cells[2].Value.ToString();
                    line3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line3.Borders.LineStyle = Excel.Constants.xlSolid;
                    line3.Font.Name = "Times New Roman";

                    Excel.Range line4 = oSheet.get_Range("D" + (line + i).ToString(), "D" + (line + i).ToString());
                    line4.Value = dgv_gv.Rows[i].Cells[3].Value.ToString();
                    line4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line4.Borders.LineStyle = Excel.Constants.xlSolid;
                    line4.Font.Name = "Times New Roman";

                    Excel.Range line5 = oSheet.get_Range("E" + (line + i).ToString(), "E" + (line + i).ToString());
                    line5.Value = dgv_gv.Rows[i].Cells[4].Value.ToString();
                    line5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line5.Borders.LineStyle = Excel.Constants.xlSolid;
                    line5.Font.Name = "Times New Roman";

                    Excel.Range line6 = oSheet.get_Range("F" + (line + i).ToString(), "F" + (line + i).ToString());
                    line6.Value = dgv_gv.Rows[i].Cells[5].Value.ToString();
                    line6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line6.Borders.LineStyle = Excel.Constants.xlSolid;
                    line6.Font.Name = "Times New Roman";

                    Excel.Range line7 = oSheet.get_Range("G" + (line + i).ToString(), "G" + (line + i).ToString());
                    line7.Value = dgv_gv.Rows[i].Cells[6].Value.ToString();
                    line7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line7.Borders.LineStyle = Excel.Constants.xlSolid;
                    line7.Font.Name = "Times New Roman";

                    Excel.Range line8 = oSheet.get_Range("H" + (line + i).ToString(), "H" + (line + i).ToString());
                    line8.Value = dgv_gv.Rows[i].Cells[7].Value.ToString();
                    line8.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line8.Borders.LineStyle = Excel.Constants.xlSolid;
                    line8.Font.Name = "Times New Roman";

                    Excel.Range line9 = oSheet.get_Range("I" + (line + i).ToString(), "I" + (line + i).ToString());
                    line9.Value = dgv_gv.Rows[i].Cells[8].Value.ToString();
                    line9.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line9.Borders.LineStyle = Excel.Constants.xlSolid;
                    line9.Font.Name = "Times New Roman";

                    Excel.Range line10 = oSheet.get_Range("J" + (line + i).ToString(), "J" + (line + i).ToString());
                    line10.Value = dgv_gv.Rows[i].Cells[9].Value.ToString();
                    line10.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line10.Borders.LineStyle = Excel.Constants.xlSolid;
                    line10.Font.Name = "Times New Roman";

                    Excel.Range line11 = oSheet.get_Range("K" + (line + i).ToString(), "K" + (line + i).ToString());
                    line11.Value = dgv_gv.Rows[i].Cells[10].Value.ToString();
                    line11.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line11.Borders.LineStyle = Excel.Constants.xlSolid;
                    line11.Font.Name = "Times New Roman";

                }


                oSheet.Name = "Giảng viên";
                oExcel.Columns.AutoFit();

                oBook.Activate();

                SaveFileDialog saveFile = new SaveFileDialog();
                if (saveFile.ShowDialog() == DialogResult.OK)
                {

                    saveFile.Filter = "Các loại tập tin (*.xlsx;*.csv;*.docx)|*.xlsx;*.csv;*.docx|Tất cả các tập tin (*.*)|*.*";
                    oBook.SaveAs(saveFile.FileName.ToLower());
                    MessageBox.Show("Xuất danh sách thành công");

                }

                oExcel.Quit();

            }
            catch
            {
                MessageBox.Show("Xuất danh sách không thành công");
            }
        }
        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
             try
            {
                
                
                {
                    if(kiemtra())
                    {
                        if(kiemTraMa(txt_ma.Text))
                        {
                            if(IsValidEmail(txt_email.Text))
                            {
                                if(kiemtraSoDT(txt_sdt.Text))
                                {
                                    string ngaysinh = dtp_ngaysinh.Value.ToString("yyyy/MM/dd");
                                    string sql = "insert into GiangVien " +
                                        "values ('" + txt_ma.Text + "',N'" + txt_hoten.Text + "','" + ngaysinh + "'," +
                                        "N'" + txt_diachi.Text + "',N'" + txt_hocvi.Text + "',N'" + txt_hocham.Text + "'" +
                                        ",N'" + cb_khoa.Text + "','" + txt_email.Text + "','" + txt_sdt.Text + "',N'" + cb_gioitinh.Text + "' ,N'"+txt_donvi.Text+"')";
                                    int up = MyClass.Update(sql);
                                   if(up>0)
                                    {
                                        MessageBox.Show("Thông tin được thêm thành công", "Thông báo");
                                        txt_ma.Clear();
                                        txt_hoten.Clear();

                                        txt_diachi.Clear();
                                        cb_khoa.Text = null;
                                        txt_email.Clear();
                                        txt_sdt.Clear();
                                        cb_gioitinh.Text = null;
                                        txt_hocvi.Clear();
                                        txt_hocham.Clear();
                                        txt_donvi.Clear();
                                        LoadDL();
                                    }
                                    else
                                    {
                                        MessageBox.Show("Thông tin được thêm không thành công", "Thông báo");

                                    }
                                        

                                    
                                    
                                }
                                else
                                {
                                    MessageBox.Show("Vui lòng nhập số điện thoại hợp lệ", "Thông báo");

                                }

                            }
                            else
                            {
                                MessageBox.Show("Vui lòng nhập đúng định dạng email !", "Thông báo");

                            }

                        }
                        else
                        {
                            MessageBox.Show("Mã giảng viên này đã có trên hệ thống!", "Thông báo");

                        }
                    }
                    else
                    {
                        MessageBox.Show("Vui lòng điền đầy đủ thông tin !", "Thông báo");

                    }
                }
            }
            catch
            {
                MessageBox.Show("Lỗi không thêm được dữ liệu", "Thông báo");

            }
        }

        private void dgv_gv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void btn_sua_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                if (dgv_gv.SelectedCells.Count > 0)
                {
                    if (kiemtra())
                    {

                        if (!kiemTraMa(txt_ma.Text))
                        {
                            if (IsValidEmail(txt_email.Text))
                            {
                                if (kiemtraSoDT(txt_sdt.Text))
                                {
                                    string ngaysinh = dtp_ngaysinh.Value.ToString("yyyy/MM/dd");
                                    string sql = @"update GiangVien set 
                                       HoTen=N'" + txt_hoten.Text + "',NgaySinh='" + ngaysinh + "',DiaChi=N'" + txt_diachi.Text + "',HocVi=N'" + txt_hocvi.Text + "',HocHam=N'" + txt_hocham.Text + "',KhoaChuQuan=N'" + cb_khoa.Text + "',Email='" + txt_email.Text + "',SoDT='" + txt_sdt.Text + "',GioiTinh=N'" + cb_gioitinh.Text + "',DonViCongTac=N'"+txt_donvi.Text+"' where MaGV = '" + txt_ma.Text + "' ";
                                     int up = MyClass.Update(sql);
                                    if(up>0)
                                    {
                                        MessageBox.Show("Thông tin được sửa thành công", "Thông báo");
                                        txt_ma.Clear();
                                        txt_hoten.Clear();

                                        txt_diachi.Clear();
                                        cb_khoa.Text = null;
                                        txt_email.Clear();
                                        txt_sdt.Clear();
                                        cb_gioitinh.Text = null;
                                        txt_hocvi.Clear();
                                        txt_hocham.Clear();
                                        txt_donvi.Clear();

                                        LoadDL();

                                    }
                                    else
                                    {
                                        MessageBox.Show("Thông tin sửa không thành công", "Thông báo");
                                    }


                                }
                                else
                                {
                                    MessageBox.Show("Vui lòng nhập số điện thoại hợp lệ", "Thông báo");

                                }

                            }
                            else
                            {
                                MessageBox.Show("Vui lòng nhập đúng định dạng email !", "Thông báo");

                            }

                        }
                        else
                        {
                            MessageBox.Show("Mã giảng viên này không có trên hệ thống!", "Thông báo");

                        }
                    }
                    else
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin! ", "Thông báo");

                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng chọn giảng viên cần sửa! ", "Thông báo");

                }
               
            }
            catch
            {
                MessageBox.Show("Lỗi không sửa được thông tin! ", "Thông báo");

            }
        }

        private void btn_xoa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            { 
            if (dgv_gv.SelectedCells.Count > 0)
            {
                    

                    if (!kiemTraMa(txt_ma.Text))
                    {
                        if (IsValidEmail(txt_email.Text))
                        {
                            if (kiemtraSoDT(txt_sdt.Text))
                            {
                                DialogResult tb = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                                if (tb == DialogResult.OK)
                                {
                                    bool XOABLDCT = xoa.XoaBLDCT(txt_ma.Text);
                                    if (XOABLDCT)
                                    {
                                        bool XOAGVHDCT = xoa.XoaGVHDCT(txt_ma.Text);
                                        if (XOAGVHDCT)
                                        {
                                            bool XOAGTGV = xoa.XoaGTGV(txt_ma.Text);
                                            if (XOAGTGV)
                                            {
                                                bool XOACTTC = xoa.XoaCTTC(txt_ma.Text);
                                                if (XOACTTC)
                                                {
                                                    bool XOABGKCT = xoa.XoaBGKCT(txt_ma.Text);
                                                    if (XOABGKCT)
                                                    {
                                                        bool XOAHD = xoa.XoaHD(txt_ma.Text);
                                                        if (XOAHD)
                                                        {
                                                            bool XOAGVDT = xoa.XoaGVDT(txt_ma.Text);
                                                            if (XOAGVDT)
                                                            {
                                                                bool XOATGBB = xoa.XoaTGBB(txt_ma.Text);
                                                                if (XOATGBB)
                                                                {
                                                                    bool XOATDHTNT = xoa.XoaTDHTNT(txt_ma.Text);
                                                                    if (XOATDHTNT)
                                                                    {
                                                                        bool xoabhtct = xoa.XoaBHTCT(txt_ma.Text);
                                                                        if (xoabhtct)
                                                                        {
                                                                            bool XOABTCCT = xoa.XoaBTCCT(txt_ma.Text);
                                                                            if (XOABTCCT)
                                                                            {
                                                                                bool XOABTCHT = xoa.XoaBTCHT(txt_ma.Text);
                                                                                if (XOABTCHT)
                                                                                {
                                                                                    bool XOAGV = xoa.XoaGiangVien(txt_ma.Text);
                                                                                    if (XOAGV)
                                                                                    {
                                                                                        MessageBox.Show("Thông tin được xóa thành công", "Thông báo");
                                                                                        txt_ma.Clear();
                                                                                        txt_hoten.Clear();

                                                                                        txt_diachi.Clear();
                                                                                        cb_khoa.Text = null;
                                                                                        txt_email.Clear();
                                                                                        txt_sdt.Clear();
                                                                                        cb_gioitinh.Text = null;
                                                                                        txt_hocvi.Clear();
                                                                                        txt_hocham.Clear();
                                                                                        txt_donvi.Clear();

                                                                                        LoadDL();
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        MessageBox.Show("Thông tin xóa không thành công", "Thông báo");
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {

                                }


                                
                                
                                    

                                
                            }
                            else
                            {
                                MessageBox.Show("Vui lòng nhập số điện thoại hợp lệ", "Thông báo");

                            }

                        }
                        else
                        {
                            MessageBox.Show("Vui lòng nhập đúng định dạng email !", "Thông báo");

                        }

                    }
                    else
                    {
                        MessageBox.Show("Mã giảng viên này không có trên hệ thống!", "Thông báo");

                    }
                
                
            }
            else
            {
                MessageBox.Show("Vui lòng chọn giảng viên cần xóa! ", "Thông báo");

            }

        }
            catch
            {
                MessageBox.Show("Lỗi không xóa được thông tin! ", "Thông báo");

            }
}

        private void btn_timkiem_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            
            
            
            if(cb_loai.Text == "")
            {
                MessageBox.Show("Vui lòng chọn khóa cần tìm ! ", "Thông báo");
            }
            else
            {
                if(txt_timkiem.Text != "")
                {
                    if(cb_loai.SelectedIndex == 0)
                    {
                        try
                        {
                            string sql = "select * from GiangVien where MaGV like '%"+txt_timkiem.Text+"%' ";
                            DataTable tb = MyClass.DocDL(sql);
                            if(tb.Rows.Count > 0)
                            {
                                dgv_gv.DataSource = tb;
                                dgv_gv.Columns[0].HeaderText = "Mã giảng viên";
                                dgv_gv.Columns[1].HeaderText = "Họ Tên";
                                dgv_gv.Columns[2].HeaderText = "Ngày Sinh";
                                dgv_gv.Columns[3].HeaderText = "Địa chỉ";
                                dgv_gv.Columns[4].HeaderText = "Học vị";
                                dgv_gv.Columns[5].HeaderText = "Học hàm";
                                dgv_gv.Columns[6].HeaderText = "Khoa chủ quản";
                                dgv_gv.Columns[7].HeaderText = "Email";
                                dgv_gv.Columns[8].HeaderText = "Số điện thoại";
                                dgv_gv.Columns[9].HeaderText = "Giới tính";
                                dgv_gv.Columns[10].HeaderText = "Đơn vị công tác";

                            }
                            else
                            {
                                MessageBox.Show("Không tìm thấy thông tin vừa nhập ! ", "Thông báo");

                            }
                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm vui lòng kiểm tra lại! ", "Thông báo");

                        }

                        return;
                    }

                    if (cb_loai.Text == "Họ Tên")
                    {
                        try
                        {
                            string sql = "select * from GiangVien where HoTen like N'%" + txt_timkiem.Text + "%' ";
                            DataTable tb = MyClass.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_gv.DataSource = tb;
                                dgv_gv.Columns[0].HeaderText = "Mã giảng viên";
                                dgv_gv.Columns[1].HeaderText = "Họ Tên";
                                dgv_gv.Columns[2].HeaderText = "Ngày Sinh";
                                dgv_gv.Columns[3].HeaderText = "Địa chỉ";
                                dgv_gv.Columns[4].HeaderText = "Học vị";
                                dgv_gv.Columns[5].HeaderText = "Học hàm";
                                dgv_gv.Columns[6].HeaderText = "Khoa chủ quản";
                                dgv_gv.Columns[7].HeaderText = "Email";
                                dgv_gv.Columns[8].HeaderText = "Số điện thoại";
                                dgv_gv.Columns[9].HeaderText = "Giới tính";
                                dgv_gv.Columns[10].HeaderText = "Đơn vị công tác";

                            }
                            else
                            {
                                MessageBox.Show("Không tìm thấy thông tin vừa nhập ! ", "Thông báo");

                            }
                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm vui lòng kiểm tra lại! ", "Thông báo");

                        }

                        return;
                    }


                    if (cb_loai.Text == "Khoa")
                    {
                        try
                        {
                            string sql = "select * from GiangVien where KhoaChuQuan like N'%" + txt_timkiem.Text + "%' ";
                            DataTable tb = MyClass.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_gv.DataSource = tb;
                                dgv_gv.Columns[0].HeaderText = "Mã giảng viên";
                                dgv_gv.Columns[1].HeaderText = "Họ Tên";
                                dgv_gv.Columns[2].HeaderText = "Ngày Sinh";
                                dgv_gv.Columns[3].HeaderText = "Địa chỉ";
                                dgv_gv.Columns[4].HeaderText = "Học vị";
                                dgv_gv.Columns[5].HeaderText = "Học hàm";
                                dgv_gv.Columns[6].HeaderText = "Khoa chủ quản";
                                dgv_gv.Columns[7].HeaderText = "Email";
                                dgv_gv.Columns[8].HeaderText = "Số điện thoại";
                                dgv_gv.Columns[9].HeaderText = "Giới tính";
                                dgv_gv.Columns[10].HeaderText = "Đơn vị công tác";

                            }
                            else
                            {
                                MessageBox.Show("Không tìm thấy thông tin vừa nhập ! ", "Thông báo");

                            }
                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm vui lòng kiểm tra lại! ", "Thông báo");

                        }

                        return;
                    }

                }
                else
                {
                    MessageBox.Show("Vui lòng nhập thông tin tìm kiếm ! ", "Thông báo");

                }
            }
        }

        private void btn_export_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ExcelExport();
        }

        private void btn_refresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                LoadDL();
                txt_timkiem.Clear();
                txt_ma.Clear();
                txt_hoten.Clear();
                txt_hocvi.Clear();
                txt_hocham.Clear();
                txt_diachi.Clear();
                txt_email.Clear();
                txt_sdt.Clear();
                txt_donvi.Clear();
                cb_khoa.SelectedIndex = -1;
                cb_loai.SelectedIndex = -1;
                cb_gioitinh.SelectedIndex = -1;
            }
            catch
            {
                MessageBox.Show("Lỗi không làm mới trang ! ", "Thông báo");
            }
        }

        private void btn_giayto_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if(!string.IsNullOrWhiteSpace(txt_ma.Text))
            {
                frm_giaytoSv fr = new frm_giaytoSv();
                fr.Magv = txt_ma.Text;
                fr.Show();
            }
            else
            {
                MessageBox.Show("Vui lòng chọn giảng viên cần xem ! ", "Thông báo");
            }
           

        }

        private void dgv_gv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_ma.Text = dgv_gv.CurrentRow.Cells[0].Value.ToString();
                txt_hoten.Text = dgv_gv.CurrentRow.Cells[1].Value.ToString();
                dtp_ngaysinh.Text = dgv_gv.CurrentRow.Cells[2].Value.ToString();
                txt_diachi.Text = dgv_gv.CurrentRow.Cells[3].Value.ToString();
                txt_hocvi.Text = dgv_gv.CurrentRow.Cells[4].Value.ToString();
                txt_hocham.Text = dgv_gv.CurrentRow.Cells[5].Value.ToString();
                cb_khoa.Text = dgv_gv.CurrentRow.Cells[6].Value.ToString();
                txt_email.Text = dgv_gv.CurrentRow.Cells[7].Value.ToString();
                txt_sdt.Text = dgv_gv.CurrentRow.Cells[8].Value.ToString();
                cb_gioitinh.Text = dgv_gv.CurrentRow.Cells[9].Value.ToString();
                txt_donvi.Text = dgv_gv.CurrentRow.Cells[10].Value.ToString();

            }
            catch
            {
                MessageBox.Show("Lỗi lấy dữ liệu từ danh sách giảng viên ", "Thông báo");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (cb_loai.Text == "")
            {
                MessageBox.Show("Vui lòng chọn khóa cần tìm ! ", "Thông báo");
            }
            else
            {
                if (txt_timkiem.Text != "")
                {
                    if (cb_loai.SelectedIndex == 0)
                    {
                        try
                        {
                            string sql = "select * from GiangVien where MaGV like '%" + txt_timkiem.Text + "%' ";
                            DataTable tb = MyClass.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_gv.DataSource = tb;
                                dgv_gv.Columns[0].HeaderText = "Mã giảng viên";
                                dgv_gv.Columns[1].HeaderText = "Họ Tên";
                                dgv_gv.Columns[2].HeaderText = "Ngày Sinh";
                                dgv_gv.Columns[3].HeaderText = "Địa chỉ";
                                dgv_gv.Columns[4].HeaderText = "Học vị";
                                dgv_gv.Columns[5].HeaderText = "Học hàm";
                                dgv_gv.Columns[6].HeaderText = "Khoa chủ quản";
                                dgv_gv.Columns[7].HeaderText = "Email";
                                dgv_gv.Columns[8].HeaderText = "Số điện thoại";
                                dgv_gv.Columns[9].HeaderText = "Giới tính";
                                dgv_gv.Columns[10].HeaderText = "Đơn vị công tác";

                            }
                            else
                            {
                                MessageBox.Show("Không tìm thấy thông tin vừa nhập ! ", "Thông báo");

                            }
                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm vui lòng kiểm tra lại! ", "Thông báo");

                        }

                        return;
                    }

                    if (cb_loai.Text == "Họ Tên")
                    {
                        try
                        {
                            string sql = "select * from GiangVien where HoTen like N'%" + txt_timkiem.Text + "%' ";
                            DataTable tb = MyClass.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_gv.DataSource = tb;
                                dgv_gv.Columns[0].HeaderText = "Mã giảng viên";
                                dgv_gv.Columns[1].HeaderText = "Họ Tên";
                                dgv_gv.Columns[2].HeaderText = "Ngày Sinh";
                                dgv_gv.Columns[3].HeaderText = "Địa chỉ";
                                dgv_gv.Columns[4].HeaderText = "Học vị";
                                dgv_gv.Columns[5].HeaderText = "Học hàm";
                                dgv_gv.Columns[6].HeaderText = "Khoa chủ quản";
                                dgv_gv.Columns[7].HeaderText = "Email";
                                dgv_gv.Columns[8].HeaderText = "Số điện thoại";
                                dgv_gv.Columns[9].HeaderText = "Giới tính";
                                dgv_gv.Columns[10].HeaderText = "Đơn vị công tác";

                            }
                            else
                            {
                                MessageBox.Show("Không tìm thấy thông tin vừa nhập ! ", "Thông báo");

                            }
                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm vui lòng kiểm tra lại! ", "Thông báo");

                        }

                        return;
                    }


                    if (cb_loai.Text == "Khoa")
                    {
                        try
                        {
                            string sql = "select * from GiangVien where KhoaChuQuan like N'%" + txt_timkiem.Text + "%' ";
                            DataTable tb = MyClass.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_gv.DataSource = tb;
                                dgv_gv.Columns[0].HeaderText = "Mã giảng viên";
                                dgv_gv.Columns[1].HeaderText = "Họ Tên";
                                dgv_gv.Columns[2].HeaderText = "Ngày Sinh";
                                dgv_gv.Columns[3].HeaderText = "Địa chỉ";
                                dgv_gv.Columns[4].HeaderText = "Học vị";
                                dgv_gv.Columns[5].HeaderText = "Học hàm";
                                dgv_gv.Columns[6].HeaderText = "Khoa chủ quản";
                                dgv_gv.Columns[7].HeaderText = "Email";
                                dgv_gv.Columns[8].HeaderText = "Số điện thoại";
                                dgv_gv.Columns[9].HeaderText = "Giới tính";
                                dgv_gv.Columns[10].HeaderText = "Đơn vị công tác";

                            }
                            else
                            {
                                MessageBox.Show("Không tìm thấy thông tin vừa nhập ! ", "Thông báo");

                            }
                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm vui lòng kiểm tra lại! ", "Thông báo");

                        }

                        return;
                    }

                }
                else
                {
                    MessageBox.Show("Vui lòng nhập thông tin tìm kiếm ! ", "Thông báo");

                }
            }
        }
    }
}
