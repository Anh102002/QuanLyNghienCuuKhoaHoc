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
using Excel = Microsoft.Office.Interop.Excel;
namespace QL_NCKH
{
    public partial class uc_capnhatketqua : DevExpress.XtraEditors.XtraUserControl
    {
        MyClass my = new MyClass();
        private List<string> productList;
        public uc_capnhatketqua()
        {
            InitializeComponent();
        }
        public void LoadDL()
        {
            try
            {
                string sql = @"select 
                            CTDeTai.MaDeTai,
                            DeTai.TenDeTai,
                            CTDeTai.NgayNghiemThu,CTDeTai.KetQuaNghiemThu,CTDeTai.SoQuyetDinh,CTDeTai.NgayThanhLapHoiDong,
                            CTDeTai.XepLoai,CTDeTai.KinhPhi,CTDeTai.GiaiThuong,CTDeTai.DanhGia
                            from CTDeTai 
                            JOIN DeTai on DeTai.MaDeTai = CTDeTai.MaDeTai ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    dgv_dt.DataSource = tb;
                    dgv_dt.Columns[0].HeaderText = "Mã đề tàì";
                    dgv_dt.Columns[1].HeaderText = "Tên đề tài";
                    dgv_dt.Columns[1].Width = 150;
                    dgv_dt.Columns[2].HeaderText = "Ngày nghiệm thu";
                    dgv_dt.Columns[3].HeaderText = "Kết quả nghiệm thu";
                    dgv_dt.Columns[4].HeaderText = "Số quyết định";
                    dgv_dt.Columns[5].HeaderText = "Ngày thành lập hội đồng";
                    dgv_dt.Columns[6].HeaderText = "Xếp loại";
                    dgv_dt.Columns[7].HeaderText = "Kinh phí";
                    dgv_dt.Columns[8].HeaderText = "Giải thưởng";
                    dgv_dt.Columns[9].HeaderText = "Đánh giá";

                }
            }
            catch
            {
                MessageBox.Show("Lỗi lấy dữ liệu vào danh sách", "Thông báo");
            }
        }
        private void uc_capnhatketqua_Load(object sender, EventArgs e)
        {
            LoadDL();
            LoadProductList();
        }

        private void dgv_dt_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_ma.Text = dgv_dt.CurrentRow.Cells[0].Value.ToString();
                
                dtp_ngaynghiemthu.Text = dgv_dt.CurrentRow.Cells[2].Value.ToString();
                txt_kq.Text = dgv_dt.CurrentRow.Cells[3].Value.ToString();
                txt_soqd.Text = dgv_dt.CurrentRow.Cells[4].Value.ToString();
                dtp_ngaythanhlaphd.Text = dgv_dt.CurrentRow.Cells[5].Value.ToString();
                txt_xeploai.Text = dgv_dt.CurrentRow.Cells[6].Value.ToString();
                txt_kinhphi.Text = dgv_dt.CurrentRow.Cells[7].Value.ToString();
                txt_giaithuong.Text = dgv_dt.CurrentRow.Cells[8].Value.ToString();
                txt_danhgia.Text = dgv_dt.CurrentRow.Cells[9].Value.ToString();

            }catch(Exception ex)
            {
                MessageBox.Show(" $ Lỗi cellclick {"+ex.Message+"}","Lỗi");
            }
        }
        public bool kiemtra()
        {
            if (string.IsNullOrWhiteSpace(txt_ma.Text) || string.IsNullOrWhiteSpace(txt_danhgia.Text)
                || string.IsNullOrWhiteSpace(txt_giaithuong.Text)
                || string.IsNullOrWhiteSpace(txt_kinhphi.Text) || string.IsNullOrWhiteSpace(txt_kq.Text)
                || string.IsNullOrWhiteSpace(txt_soqd.Text) || string.IsNullOrWhiteSpace(txt_xeploai.Text))
            {
                return false;
            }

            return true;
        }
        public bool kiemTraMa(string ma)
        {
            try
            {
                string sql = "select * from CTDeTai Where MaDeTai = '" + ma + "' ";
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
        public bool kiemTraMaDT(string ma)
        {
            try
            {
                string sql = "select * from DeTai Where MaDeTai = '" + ma + "' ";
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
        public bool KtraKinhPhi()
        {
            int count = 0;
            string kinhphi = txt_kinhphi.Text;
            for (int i = 0; i < kinhphi.Length; i++)
            {
                char kp = kinhphi[i];
                if(char.IsLetter(kp))
                {
                    count++;
                }
            }

            if(count > 0)
            {
                return false;
            }
            return true;
        }
        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {



                if (kiemtra())
                {
                    if(!kiemTraMaDT(txt_ma.Text))
                    {
                        if (kiemTraMa(txt_ma.Text))
                        {
                            if (KtraKinhPhi())
                            {
                                string ngaynt = dtp_ngaynghiemthu.Value.ToString("yyyy/MM/dd");
                                string ngaytlhd = dtp_ngaythanhlaphd.Value.ToString("yyyy/MM/dd");
                                string sql = "insert into CTDeTai values('" + txt_ma.Text + "','" + ngaynt + "',N'" + txt_kq.Text + "',N'" + txt_soqd.Text + "','" + ngaytlhd + "',N'" + txt_xeploai.Text + "','" + txt_kinhphi.Text + "',N'" + txt_giaithuong.Text + "',N'" + txt_danhgia.Text + "')";
                                int up = my.Update(sql);
                                if (up > 0)
                                {
                                    MessageBox.Show("Thông tin được thêm thành công", "Thông báo");
                                    txt_ma.Clear();
                                    
                                    txt_kq.Clear();
                                    txt_soqd.Clear();
                                    txt_xeploai.Clear();
                                    txt_kinhphi.Clear();
                                    txt_tendetai.Clear();
                                    txt_giaithuong.Clear();
                                    txt_danhgia.Clear();

                                    LoadDL();
                                }
                                else
                                {
                                    MessageBox.Show("Thông tin thêm không thành công", "Thông báo");
                                }
                            }
                            else
                            {
                                MessageBox.Show("Vui lòng nhập kinh phí là số", "Thông báo");
                                txt_kinhphi.Clear();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Mã đề tài đã tồn tại !! ", "Thông báo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Hãy nhập mã đề tài có trong hệ thống!! ", "Thông báo");
                    }

                }
                else
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin !!", "Thông báo");
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
                
                    if (kiemtra())
                    {

                        if (!kiemTraMa(txt_ma.Text))
                        {
                            if (KtraKinhPhi())
                            {
                                
                                    string ngaynt = dtp_ngaynghiemthu.Value.ToString("yyyy/MM/dd");
                                    string ngaytlhd = dtp_ngaythanhlaphd.Value.ToString("yyyy/MM/dd");
                                    string sql = "update CTDeTai set NgayNghiemThu = '" + ngaynt + "',KetQuaNghiemThu=N'" + txt_kq.Text + "',SoQuyetDinh=N'" + txt_soqd.Text + "',NgayThanhLapHoiDong='" + ngaytlhd + "',XepLoai=N'" + txt_xeploai.Text + "',KinhPhi='" + txt_kinhphi.Text + "',GiaiThuong=N'" + txt_giaithuong.Text + "',DanhGia=N'" + txt_danhgia.Text + "' where  MaDeTai='" + txt_ma.Text + "' ";
                                    int up = my.Update(sql);
                                    if (up > 0)
                                    {
                                        MessageBox.Show("Thông tin được sửa thành công", "Thông báo");
                                    txt_ma.Clear();
                                    
                                    txt_kq.Clear();
                                    txt_soqd.Clear();
                                    txt_xeploai.Clear();
                                    txt_kinhphi.Clear();
                                    txt_giaithuong.Clear();
                                txt_tendetai.Clear();
                                txt_danhgia.Clear();

                                LoadDL();
                                    }
                                    else
                                    {
                                        MessageBox.Show("Thông tin sửa không thành công", "Thông báo");

                                    }



                                

                            }
                            else
                            {
                                MessageBox.Show("Vui lòng nhập kinh phí là số !", "Thông báo");

                            }

                        }
                        else
                        {
                            MessageBox.Show("Mã đề tài này không có trên hệ thống!", "Thông báo");

                        }
                    }
                    else
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin! ", "Thông báo");

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
                if (kiemtra())
                {

                    if (!kiemTraMa(txt_ma.Text))
                    {
                        if (KtraKinhPhi())
                        {

                            DialogResult result = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                            if (result == DialogResult.OK)
                            {
                                string sql = "delete from CTDeTai where MaDeTai = '" + txt_ma.Text + "' ";
                                int up = my.Update(sql);
                                if (up > 0)
                                {
                                    MessageBox.Show("Thông tin được xóa thành công", "Thông báo");
                                    txt_ma.Clear();

                                    txt_kq.Clear();
                                    txt_soqd.Clear();
                                    txt_xeploai.Clear();
                                    txt_kinhphi.Clear();
                                    txt_giaithuong.Clear();
                                    txt_danhgia.Clear();
                                    txt_tendetai.Clear();
                                    LoadDL();


                                }
                                else
                                {
                                    MessageBox.Show("Thông tin xóa không thành công", "Thông báo");
                                }
                            }
                            else
                            {

                            }
                                



                            

                        }
                        else
                        {
                            MessageBox.Show("Vui lòng nhập kinh phí là số !", "Thông báo");

                        }

                    }
                    else
                    {
                        MessageBox.Show("Mã đề tài này không có trên hệ thống!", "Thông báo");

                    }


                }
                else
                {
                    MessageBox.Show("Vui lòng chọn đề tài cần xóa! ", "Thông báo");

                }

            }
            catch
            {
                MessageBox.Show("Lỗi không xóa được thông tin! ", "Thông báo");

            }
        }

        private void btn_timkiem_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if(!string.IsNullOrWhiteSpace(txt_timkiem.Text))
            {
                int i = cbo_tk.SelectedIndex;
                string sql;
                
                switch (i)
                {
                    case 0:
                        sql= @"select 
                            CTDeTai.MaDeTai,
                            DeTai.TenDeTai,
                            CTDeTai.NgayNghiemThu,CTDeTai.KetQuaNghiemThu,CTDeTai.SoQuyetDinh,CTDeTai.NgayThanhLapHoiDong,
                            CTDeTai.XepLoai,CTDeTai.KinhPhi,CTDeTai.GiaiThuong,CTDeTai.DanhGia
                            from CTDeTai 
                            JOIN DeTai on DeTai.MaDeTai = CTDeTai.MaDeTai where CTDeTai.MaDeTai like '%" + txt_timkiem.Text + "%' ";
                        DataTable tb = my.DocDL(sql);
                        if (tb.Rows.Count > 0)
                        {
                            dgv_dt.DataSource = tb;
                            dgv_dt.Columns[0].HeaderText = "Mã đề tàì";
                            dgv_dt.Columns[1].HeaderText = "Tên đề tài";
                            dgv_dt.Columns[1].Width = 150;
                            dgv_dt.Columns[2].HeaderText = "Ngày nghiệm thu";
                            dgv_dt.Columns[3].HeaderText = "Kết quả nghiệm thu";
                            dgv_dt.Columns[4].HeaderText = "Số quyết định";
                            dgv_dt.Columns[5].HeaderText = "Ngày thành lập hội đồng";
                            dgv_dt.Columns[6].HeaderText = "Xếp loại";
                            dgv_dt.Columns[7].HeaderText = "Kinh phí";
                            dgv_dt.Columns[8].HeaderText = "Giải thưởng";
                            dgv_dt.Columns[9].HeaderText = "Đánh giá";

                        }
                        break;
                    case 1:
                        sql = @"select 
                            CTDeTai.MaDeTai,
                            DeTai.TenDeTai,
                            CTDeTai.NgayNghiemThu,CTDeTai.KetQuaNghiemThu,CTDeTai.SoQuyetDinh,CTDeTai.NgayThanhLapHoiDong,
                            CTDeTai.XepLoai,CTDeTai.KinhPhi,CTDeTai.GiaiThuong,CTDeTai.DanhGia
                            from CTDeTai 
                            JOIN DeTai on DeTai.MaDeTai = CTDeTai.MaDeTai where CTDeTai.SoQuyetDinh like N'%" + txt_timkiem.Text + "%' ";
                        DataTable tb1 = my.DocDL(sql);
                        if (tb1.Rows.Count > 0)
                        {
                            dgv_dt.DataSource = tb1;
                            dgv_dt.Columns[0].HeaderText = "Mã đề tàì";
                            dgv_dt.Columns[1].HeaderText = "Tên đề tài";
                            dgv_dt.Columns[1].Width = 150;
                            dgv_dt.Columns[2].HeaderText = "Ngày nghiệm thu";
                            dgv_dt.Columns[3].HeaderText = "Kết quả nghiệm thu";
                            dgv_dt.Columns[4].HeaderText = "Số quyết định";
                            dgv_dt.Columns[5].HeaderText = "Ngày thành lập hội đồng";
                            dgv_dt.Columns[6].HeaderText = "Xếp loại";
                            dgv_dt.Columns[7].HeaderText = "Kinh phí";
                            dgv_dt.Columns[8].HeaderText = "Giải thưởng";
                            dgv_dt.Columns[9].HeaderText = "Đánh giá";

                        }
                        break;
                    case 2:
                        sql = @"select 
                            CTDeTai.MaDeTai,
                            DeTai.TenDeTai,
                            CTDeTai.NgayNghiemThu,CTDeTai.KetQuaNghiemThu,CTDeTai.SoQuyetDinh,CTDeTai.NgayThanhLapHoiDong,
                            CTDeTai.XepLoai,CTDeTai.KinhPhi,CTDeTai.GiaiThuong,CTDeTai.DanhGia
                            from CTDeTai 
                            JOIN DeTai on DeTai.MaDeTai = CTDeTai.MaDeTai where CTDeTai.XepLoai like N'%" + txt_timkiem.Text + "%' ";
                        DataTable tb2 = my.DocDL(sql);
                        if (tb2.Rows.Count > 0)
                        {
                            dgv_dt.DataSource = tb2;
                            dgv_dt.Columns[0].HeaderText = "Mã đề tàì";
                            dgv_dt.Columns[1].HeaderText = "Tên đề tài";
                            dgv_dt.Columns[1].Width = 150;
                            dgv_dt.Columns[2].HeaderText = "Ngày nghiệm thu";
                            dgv_dt.Columns[3].HeaderText = "Kết quả nghiệm thu";
                            dgv_dt.Columns[4].HeaderText = "Số quyết định";
                            dgv_dt.Columns[5].HeaderText = "Ngày thành lập hội đồng";
                            dgv_dt.Columns[6].HeaderText = "Xếp loại";
                            dgv_dt.Columns[7].HeaderText = "Kinh phí";
                            dgv_dt.Columns[8].HeaderText = "Giải thưởng";
                            dgv_dt.Columns[9].HeaderText = "Đánh giá";

                        }
                        break;

                }
            }
            else
            {
                MessageBox.Show("Vui lòng nhập thông tin tìm kiếm ","Thông báo");
            }
            
            
            

        }
        public void ExcelExport()
        {
            try
            {

                string sql = @"select 
                            CTDeTai.MaDeTai,
                            DeTai.TenDeTai,
                            CTDeTai.NgayNghiemThu,CTDeTai.KetQuaNghiemThu,CTDeTai.SoQuyetDinh,CTDeTai.NgayThanhLapHoiDong,
                            CTDeTai.XepLoai,CTDeTai.KinhPhi,CTDeTai.GiaiThuong,CTDeTai.DanhGia
                            from CTDeTai 
                            JOIN DeTai on DeTai.MaDeTai = CTDeTai.MaDeTai ";
                DataTable tb = my.DocDL(sql);

                Excel.Application oExcel = new Excel.Application();
                Excel.Workbook oBook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

                Excel.Worksheet oSheet = (Excel.Worksheet)oBook.Worksheets[1];

                Excel.Range head = oSheet.get_Range("A1", "I1");

                head.MergeCells = true;

                head.Value2 = "KẾT QUẢ NGHIỆM THU";

                head.Font.Bold = true;

                head.Font.Name = "Times New Roman";

                head.Font.Size = "20";

                head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range cl1 = oSheet.get_Range("A3", "A3");
                cl1.Value = "Mã đề tài";

                Excel.Range cl2 = oSheet.get_Range("B3", "B3");
                cl2.Value = "Tên đề tài";
                Excel.Range cl3 = oSheet.get_Range("C3", "C3");
                cl3.Value = "Ngày nghiệm thu";

                Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                cl4.Value = "Kết Qủa nghiệm thu";

                Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                cl5.Value = "Số quyết định";

                Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                cl6.Value = "Ngày thành lập hội đồng";
                Excel.Range cl7 = oSheet.get_Range("G3", "G3");
                cl7.Value = "Xếp loại";

                Excel.Range cl8 = oSheet.get_Range("H3", "H3");
                cl8.Value = "Kinh phí";

                Excel.Range cl9 = oSheet.get_Range("I3", "I3");
                cl9.Value = "Giải thưởng";

                Excel.Range cl10 = oSheet.get_Range("J3", "J3");
                cl10.Value = "Đánh giá";



                Excel.Range rowHead = oSheet.get_Range("A3", "J3");
                rowHead.Font.Bold = true;
                rowHead.Font.Size = 13;
                rowHead.Borders.LineStyle = Excel.Constants.xlSolid;
                rowHead.Interior.ColorIndex = 6;
                rowHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int line = 4;
                for (int i = 0; i < tb.Rows.Count ; i++)
                {
                    Excel.Range line1 = oSheet.get_Range("A" + (line + i).ToString(), "A" + (line + i).ToString());
                    line1.Value = tb.Rows[i][0].ToString();
                    line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line1.Borders.LineStyle = Excel.Constants.xlSolid;

                    Excel.Range line2 = oSheet.get_Range("B" + (line + i).ToString(), "B" + (line + i).ToString());
                    line2.Value = tb.Rows[i][1].ToString();
                    line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line2.Borders.LineStyle = Excel.Constants.xlSolid;


                    Excel.Range line3 = oSheet.get_Range("C" + (line + i).ToString(), "C" + (line + i).ToString());
                    line3.Value = tb.Rows[i][2].ToString();
                    line3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line3.Borders.LineStyle = Excel.Constants.xlSolid;

                    Excel.Range line4 = oSheet.get_Range("D" + (line + i).ToString(), "D" + (line + i).ToString());
                    line4.Value = tb.Rows[i][3].ToString();
                    line4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line4.Borders.LineStyle = Excel.Constants.xlSolid;

                    Excel.Range line5 = oSheet.get_Range("E" + (line + i).ToString(), "E" + (line + i).ToString());
                    line5.Value = tb.Rows[i][4].ToString();
                    line5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line5.Borders.LineStyle = Excel.Constants.xlSolid;


                    Excel.Range line6 = oSheet.get_Range("F" + (line + i).ToString(), "F" + (line + i).ToString());
                    line6.Value = tb.Rows[i][5].ToString();
                    line6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line6.Borders.LineStyle = Excel.Constants.xlSolid;


                    Excel.Range line7 = oSheet.get_Range("G" + (line + i).ToString(), "G" + (line + i).ToString());
                    line7.Value = tb.Rows[i][6].ToString();
                    line7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line7.Borders.LineStyle = Excel.Constants.xlSolid;

                    Excel.Range line8 = oSheet.get_Range("H" + (line + i).ToString(), "H" + (line + i).ToString());
                    line8.Value = tb.Rows[i][7].ToString();
                    line8.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line8.Borders.LineStyle = Excel.Constants.xlSolid;

                    Excel.Range line9 = oSheet.get_Range("I" + (line + i).ToString(), "I" + (line + i).ToString());
                    line9.Value = tb.Rows[i][8].ToString();
                    line9.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line9.Borders.LineStyle = Excel.Constants.xlSolid;

                    Excel.Range line10 = oSheet.get_Range("J" + (line + i).ToString(), "J" + (line + i).ToString());
                    line10.Value = tb.Rows[i][9].ToString();
                    line10.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line10.Borders.LineStyle = Excel.Constants.xlSolid;


                }


                oSheet.Name = "KQNT";
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
        private void btn_export_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ExcelExport();
        }

        private void btn_refesh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadDL();
            txt_danhgia.Clear();
            txt_giaithuong.Clear();
            txt_kinhphi.Clear();
            txt_kq.Clear();
            txt_ma.Clear();
            txt_soqd.Clear();
            txt_timkiem.Clear();
            txt_tendetai.Clear();
            txt_xeploai.Clear();
            cbo_tk.SelectedIndex = -1;
        }

        private void btn_baocao_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_ma.Text))
            {
                MessageBox.Show("Vui lòng chọn đề tài muốn xem báo cáo !", "Thông báo");
            }
            else
            {
                frm_filebaocao frm = new frm_filebaocao();
                frm.Madetai = txt_ma.Text;
                frm.Show();
            }
        }
        private void LoadProductList()
        {
            try
            {

                productList = new List<string>();
                string query = "SELECT MaDeTai FROM DeTai";
                DataTable tb = my.DocDL(query);
                if (tb.Rows.Count > 0)
                {
                    for (int i = 0; i < tb.Rows.Count; i++)
                    {
                        string ma = tb.Rows[i][0].ToString();
                        productList.Add(ma);
                    }
                }
            }
            catch
            {
                MessageBox.Show($"Lỗi thực hiện tạo danh sách đề tài", "Lỗi");
            }

        }
        private void ShowSuggestions(List<string> suggestions)
        {
            list_detai.Items.Clear();
            list_detai.Items.AddRange(suggestions.ToArray());

            list_detai.Visible = suggestions.Any();
        }
        private void txt_ma_TextChanged(object sender, EventArgs e)
        {
            
                string searchTerm = txt_ma.Text.ToLower();
                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
                    List<string> filteredProducts = productList
                   .Where(product => product.ToLower().Contains(searchTerm))
                   .ToList();

                    ShowSuggestions(filteredProducts);
                }
                else
                {
                list_detai.Visible = false;
                }


            
        }

        private void list_detai_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_detai.SelectedItem != null)
            {
                string selectedProduct = list_detai.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_ma.Text = selectedProduct;
                    list_detai.Visible = false;
                    string sql = "select TenDeTai from DeTai where MaDeTai = '" + selectedProduct + "' ";
                    DataTable tb = my.DocDL(sql);
                    if (tb.Rows.Count > 0)
                    { 
                        string hoten = tb.Rows[0][0].ToString();
                        txt_tendetai.Text = hoten;
                    }
                }

            }
        }
    }
}
