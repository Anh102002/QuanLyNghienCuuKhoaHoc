using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
namespace QL_NCKH
{
    public partial class uc_sinhvienngoaitruong : DevExpress.XtraEditors.XtraUserControl
    {
        public uc_sinhvienngoaitruong()
        {
            InitializeComponent();
        }
        MyClass my = new MyClass();

        public void LoadDL()
        {
            try
            {
                string sql = "select * from SinhVienNgoaiTruong";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    dgv_sinhvien.DataSource = tb;
                    dgv_sinhvien.Columns[0].HeaderText = "Mã sinh viên";
                    dgv_sinhvien.Columns[1].HeaderText = "Họ tên";
                    dgv_sinhvien.Columns[1].Width = 150;
                    dgv_sinhvien.Columns[2].HeaderText = "Ngày sinh";
                    dgv_sinhvien.Columns[3].HeaderText = "Email";
                    dgv_sinhvien.Columns[4].HeaderText = "Số điện thoại";
                    dgv_sinhvien.Columns[5].HeaderText = "Giới tính";
                    dgv_sinhvien.Columns[6].Width = 250;
                    dgv_sinhvien.Columns[6].HeaderText = "Đơn vị";
                    
                    

                }
            }
            catch
            {
                MessageBox.Show("Lỗi lấy dữ liệu vào danh sách", "Thông báo");
            }
        }
        private void uc_sinhvienngoaitruong_Load(object sender, EventArgs e)
        {
            LoadDL();
        }
        public bool kiemtra()
        {
            if (string.IsNullOrWhiteSpace(txt_masv.Text) || string.IsNullOrWhiteSpace(txt_hoten.Text)                
                || string.IsNullOrWhiteSpace(txt_email.Text) || string.IsNullOrWhiteSpace(txt_donvi.Text)
                || string.IsNullOrWhiteSpace(cb_gioitinh.Text) || string.IsNullOrWhiteSpace(txt_sodt.Text))
            {
                return false;
            }

            return true;
        }
        public bool kiemTraMa(string ma)
        {
            try
            {
                string sql = "select * from SinhVienNgoaiTruong Where MaSVNT = '" + ma + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
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
            sodt = sodt.Replace(" ", "");
            List<string> dienthoai = new List<string> { "03", "08", "09", "07", "05" };
            int count = 0;
            foreach (var item in sodt)
            {
                if (!char.IsDigit(item))
                {
                    count++;
                }
            }
            if (count == 0 && sodt.Count() < 10)
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

                Excel.Range head = oSheet.get_Range("A1", "G1");

                head.MergeCells = true;

                head.Value2 = "DANH SÁCH SINH VIÊN NGOÀI TRƯỜNG";

                head.Font.Bold = true;

                head.Font.Name = "Times New Roman";

                head.Font.Size = "20";

                head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range cl1 = oSheet.get_Range("A3", "A3");
                cl1.Value = "Mã sinh viên";

                Excel.Range cl2 = oSheet.get_Range("B3", "B3");
                cl2.Value = "Họ Tên";
                Excel.Range cl3 = oSheet.get_Range("C3", "C3");
                cl3.Value = "Ngày sinh";

                Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                cl4.Value = "Email";

                Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                cl5.Value = "Số điện thoại";

                Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                cl6.Value = "Giới tính";
                Excel.Range cl7 = oSheet.get_Range("G3", "G3");
                cl7.Value = "Đơn vị";

                //Excel.Range cl8 = oSheet.get_Range("H3", "H3");
                //cl8.Value = "Cở sở";

                //Excel.Range cl9 = oSheet.get_Range("I3", "I3");
                //cl9.Value = "Giới tính";



                Excel.Range rowHead = oSheet.get_Range("A3", "G3");
                rowHead.Font.Bold = true;
                rowHead.Font.Size = 13;
                rowHead.Font.Name = "Times New Roman";
                rowHead.Borders.LineStyle = Excel.Constants.xlSolid;
                rowHead.Interior.ColorIndex = 6;
                rowHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int line = 4;
                for (int i = 0; i < dgv_sinhvien.Rows.Count - 1; i++)
                {
                    Excel.Range line1 = oSheet.get_Range("A" + (line + i).ToString(), "A" + (line + i).ToString());
                    line1.Value = dgv_sinhvien.Rows[i].Cells[0].Value.ToString();
                    line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line1.Borders.LineStyle = Excel.Constants.xlSolid;
                    line1.Font.Name = "Times New Roman";

                    Excel.Range line2 = oSheet.get_Range("B" + (line + i).ToString(), "B" + (line + i).ToString());
                    line2.Value = dgv_sinhvien.Rows[i].Cells[1].Value.ToString();
                    line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line2.Borders.LineStyle = Excel.Constants.xlSolid;
                    line2.Font.Name = "Times New Roman";

                    Excel.Range line3 = oSheet.get_Range("C" + (line + i).ToString(), "C" + (line + i).ToString());
                    line3.Value = dgv_sinhvien.Rows[i].Cells[2].Value.ToString();
                    line3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line3.Borders.LineStyle = Excel.Constants.xlSolid;
                    line3.Font.Name = "Times New Roman";

                    Excel.Range line4 = oSheet.get_Range("D" + (line + i).ToString(), "D" + (line + i).ToString());
                    line4.Value = dgv_sinhvien.Rows[i].Cells[3].Value.ToString();
                    line4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line4.Borders.LineStyle = Excel.Constants.xlSolid;
                    line4.Font.Name = "Times New Roman";

                    Excel.Range line5 = oSheet.get_Range("E" + (line + i).ToString(), "E" + (line + i).ToString());
                    line5.Value = dgv_sinhvien.Rows[i].Cells[4].Value.ToString();
                    line5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line5.Borders.LineStyle = Excel.Constants.xlSolid;
                    line5.Font.Name = "Times New Roman";

                    Excel.Range line6 = oSheet.get_Range("F" + (line + i).ToString(), "F" + (line + i).ToString());
                    line6.Value = dgv_sinhvien.Rows[i].Cells[5].Value.ToString();
                    line6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line6.Borders.LineStyle = Excel.Constants.xlSolid;
                    line6.Font.Name = "Times New Roman";

                    Excel.Range line7 = oSheet.get_Range("G" + (line + i).ToString(), "G" + (line + i).ToString());
                    line7.Value = dgv_sinhvien.Rows[i].Cells[6].Value.ToString();
                    line7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line7.Borders.LineStyle = Excel.Constants.xlSolid;
                    line7.Font.Name = "Times New Roman";

                    //Excel.Range line8 = oSheet.get_Range("H" + (line + i).ToString(), "H" + (line + i).ToString());
                    //line8.Value = dgv_sinhvien.Rows[i].Cells[7].Value.ToString();
                    //line8.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    //line8.Borders.LineStyle = Excel.Constants.xlSolid;
                    //line8.Font.Name = "Times New Roman";

                    //Excel.Range line9 = oSheet.get_Range("I" + (line + i).ToString(), "I" + (line + i).ToString());
                    //line9.Value = dgv_sinhvien.Rows[i].Cells[8].Value.ToString();
                    //line9.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    //line9.Borders.LineStyle = Excel.Constants.xlSolid;
                    //line9.Font.Name = "Times New Roman";


                }


                oSheet.Name = "SVNT";
                oExcel.Columns.AutoFit();

                oBook.Activate();

                SaveFileDialog saveFile = new SaveFileDialog();
                if (saveFile.ShowDialog() == DialogResult.OK)
                {

                    saveFile.Filter = "Text files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
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



                if (kiemtra())
                {
                    if (kiemTraMa(txt_masv.Text))
                    {
                        if (IsValidEmail(txt_email.Text))
                        {
                            if (kiemtraSoDT(txt_sodt.Text))
                            {
                                string ngaysinh = dtp_ngaysinh.Value.ToString("yyyy/MM/dd");
                                string sql = "insert into SinhVienNgoaiTruong values('" + txt_masv.Text + "',N'" + txt_hoten.Text + "','" + ngaysinh + "','" + txt_email.Text + "','" + txt_sodt.Text + "',N'" + cb_gioitinh.Text + "',N'" + txt_donvi.Text + " ')";
                                int up = my.Update(sql);
                                if (up > 0)
                                {
                                    MessageBox.Show("Thông tin được thêm thành công", "Thông báo");
                                    txt_masv.Clear();
                                    txt_hoten.Clear();
                                    txt_donvi.Clear();
                                    
                                    txt_email.Clear();
                                    txt_sodt.Clear();
                                    cb_gioitinh.SelectedIndex = -1;
                                   

                                    LoadDL();
                                }
                                else
                                {
                                    MessageBox.Show("Thông tin thêm không thành công", "Thông báo");
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
            catch
            {
                MessageBox.Show("Lỗi không thêm được dữ liệu", "Thông báo");

            }
        }

        private void btn_sua_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                if (dgv_sinhvien.SelectedCells.Count > 0)
                {
                    if (kiemtra())
                    {

                        if (!kiemTraMa(txt_masv.Text))
                        {
                            if (IsValidEmail(txt_email.Text))
                            {
                                if (kiemtraSoDT(txt_sodt.Text))
                                {
                                    string ngaysinh = dtp_ngaysinh.Value.ToString("yyyy/MM/dd");
                                    string sql = "update SinhVienNgoaiTruong set HoTen = N'" + txt_hoten.Text + "',NgaySinh='" + ngaysinh + "',Email='" + txt_email.Text + "',SoDT='" + txt_sodt.Text + "',GioiTinh=N'" + cb_gioitinh.Text + "',DonVi=N'" + txt_donvi.Text + "' where  MaSVNT='" + txt_masv.Text + "' ";
                                    int up = my.Update(sql);
                                    if (up > 0)
                                    {
                                        MessageBox.Show("Thông tin được sửa thành công", "Thông báo");
                                        txt_masv.Clear();
                                        txt_hoten.Clear();
                                        txt_donvi.Clear();
                                        
                                        txt_email.Clear();
                                        txt_sodt.Clear();
                                        cb_gioitinh.SelectedIndex = -1;
                                        

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
                            MessageBox.Show("Mã sinh viên này không có trên hệ thống!", "Thông báo");

                        }
                    }
                    else
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin! ", "Thông báo");

                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng chọn sinh viên cần sửa! ", "Thông báo");

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
                if (dgv_sinhvien.SelectedCells.Count > 0)
                {
                    if (kiemtra())
                    {

                        if (!kiemTraMa(txt_masv.Text))
                        {
                            if (IsValidEmail(txt_email.Text))
                            {
                                if (kiemtraSoDT(txt_sodt.Text))
                                {

                                    DialogResult tb = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                                    if (tb == DialogResult.OK)
                                    {
                                        string svnt = "delete from ThanhVienCuocThiNT where MaSVNT='" + txt_masv.Text + "' ";
                                        int upSVNT = my.Update(svnt);
                                        if (upSVNT >= 0)
                                        {
                                            string sql = "delete from SinhVienNgoaiTruong where MaSVNT='" + txt_masv.Text + "' ";
                                            int up = my.Update(sql);
                                            if (up > 0)
                                            {
                                                MessageBox.Show("Thông tin được xóa thành công", "Thông báo");
                                                txt_masv.Clear();
                                                txt_hoten.Clear();
                                                txt_donvi.Clear();

                                                txt_email.Clear();
                                                txt_sodt.Clear();
                                                cb_gioitinh.SelectedIndex = -1;


                                                LoadDL();
                                            }
                                            else
                                            {
                                                MessageBox.Show("Thông tin xóa không thành công", "Thông báo");

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
                            MessageBox.Show("Mã sinh viên này không có trên hệ thống!", "Thông báo");

                        }
                    }
                    else
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin! ", "Thông báo");

                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng chọn sinh viên cần sửa! ", "Thông báo");

                }

            }
            catch
            {
                MessageBox.Show("Lỗi không xóa được thông tin! ", "Thông báo");

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(cbo_tk.Text))
            {
                MessageBox.Show("Vui nhập chọn khóa tìm kiếm  !", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(txt_timkiem.Text))
                {
                    MessageBox.Show("Vui nhập thông tin tìm kiếm  !", "Thông báo");
                }
                else
                {
                    if (cbo_tk.Text == "Mã sinh viên")
                    {
                        try
                        {

                            string sql = "select * from SinhVienNgoaiTruong where MaSVNT like '%"+txt_timkiem.Text+"%' ";
                            DataTable tb = my.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_sinhvien.DataSource = tb;
                                dgv_sinhvien.Columns[0].HeaderText = "Mã sinh viên";
                                dgv_sinhvien.Columns[1].HeaderText = "Họ tên";
                                dgv_sinhvien.Columns[1].Width = 150;
                                dgv_sinhvien.Columns[2].HeaderText = "Ngày sinh";
                                dgv_sinhvien.Columns[3].HeaderText = "Email";
                                dgv_sinhvien.Columns[4].HeaderText = "Số điện thoại";
                                dgv_sinhvien.Columns[5].HeaderText = "Giới tính";
                                dgv_sinhvien.Columns[6].Width = 250;
                                dgv_sinhvien.Columns[6].HeaderText = "Đơn vị";



                            }

                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm theo mã sinh viên  !", "Thông báo");
                        }
                    }
                    else if (cbo_tk.Text == "Họ tên")
                    {
                        try
                        {
                            string sql = "select * from SinhVienNgoaiTruong where HoTen like '%" + txt_timkiem.Text + "%' ";
                            DataTable tb = my.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_sinhvien.DataSource = tb;
                                dgv_sinhvien.Columns[0].HeaderText = "Mã sinh viên";
                                dgv_sinhvien.Columns[1].HeaderText = "Họ tên";
                                dgv_sinhvien.Columns[1].Width = 150;
                                dgv_sinhvien.Columns[2].HeaderText = "Ngày sinh";
                                dgv_sinhvien.Columns[3].HeaderText = "Email";
                                dgv_sinhvien.Columns[4].HeaderText = "Số điện thoại";
                                dgv_sinhvien.Columns[5].HeaderText = "Giới tính";
                                dgv_sinhvien.Columns[6].Width = 250;
                                dgv_sinhvien.Columns[6].HeaderText = "Đơn vị";



                            }

                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm theo họ tên sinh viên  !", "Thông báo");
                        }
                    }
                    

                }
            }
        }

        private void btn_refesh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                LoadDL();
                txt_masv.Clear();
                txt_hoten.Clear();
                txt_donvi.Clear();
                
                txt_email.Clear();
                txt_sodt.Clear();
                cb_gioitinh.SelectedIndex = -1;
                
            }
            catch
            {
                MessageBox.Show("Lỗi không làm mới trang", "Thông báo");
            }
        }

        private void btn_export_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ExcelExport();
        }

        private void btn_giayto_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

        }

        private void dgv_sinhvien_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txt_masv.Text = dgv_sinhvien.CurrentRow.Cells[0].Value.ToString();
            txt_hoten.Text = dgv_sinhvien.CurrentRow.Cells[1].Value.ToString();
            dtp_ngaysinh.Text = dgv_sinhvien.CurrentRow.Cells[2].Value.ToString();
            
            
            txt_email.Text = dgv_sinhvien.CurrentRow.Cells[3].Value.ToString();
            txt_sodt.Text = dgv_sinhvien.CurrentRow.Cells[4].Value.ToString();
            
            cb_gioitinh.Text = dgv_sinhvien.CurrentRow.Cells[5].Value.ToString();
            txt_donvi.Text = dgv_sinhvien.CurrentRow.Cells[6].Value.ToString();
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(txt_masv.Text))
            {
                frm_giayto1 fr = new frm_giayto1();
                string masv = txt_masv.Text;
                fr.Masv = masv;
                fr.Show();
            }
            else
            {
                MessageBox.Show("Vui lòng chọn sinh viên cần xem !!", "Thông báo");
            }
        }
    }
}
