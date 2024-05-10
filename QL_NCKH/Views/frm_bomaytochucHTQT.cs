using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
namespace QL_NCKH
{
    public partial class frm_bomaytochucHTQT : DevExpress.XtraEditors.XtraForm
    {

        public frm_bomaytochucHTQT()
        {
            InitializeComponent();
        }
        MyClass my = new MyClass();
        private string maht;
        List<string> productListKM;
        List<string> productListGV;
        public string Maht
        {
            get { return this.maht; }
            set { this.maht = value; }
        }
        public void loadDLBTC(string ma)
        {
            try
            {

                string query = @" select BanToChucHT.MaGV,GiangVien.HoTen,BanToChucHT.ChucVu,BanToChucHT.VaiTro from BanToChucHT,GiangVien
                                        WHERE BanToChucHT.MaGV = GiangVien.MaGV and BanToChucHT.MaHT = '" + ma + "'  ";
                DataTable dt = my.DocDL(query);
                dgv_nt.DataSource = dt;
                dgv_nt.Columns[0].HeaderText = "Mã giảng viên";
                dgv_nt.Columns[1].HeaderText = "Tên giảng viên";
                dgv_nt.Columns[1].Width = 150;
                dgv_nt.Columns[2].HeaderText = "Chức vụ";
                dgv_nt.Columns[3].HeaderText = "Vai trò";





            }
            catch(Exception ex)
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu ban tổ chức hội thảo  {"+ex.Message+"}", "Lỗi");
            }
        }
        public void loadDLBTCNT(string ma)
        {
            try
            {

                string query = @" select BanToChucHTNT.MaKM,TVNgoaiTruong.HoTen,BanToChucHTNT.ChucVu,BanToChucHTNT.VaiTro from BanToChucHTNT,TVNgoaiTruong
                                        WHERE BanToChucHTNT.MaKM = TVNgoaiTruong.MaKM and BanToChucHTNT.MaHT = '" + ma + "'  ";
                DataTable dt = my.DocDL(query);
                dgv_km.DataSource = dt;
                dgv_km.Columns[0].HeaderText = "Mã cán bộ ";
                dgv_km.Columns[1].HeaderText = "Tên cán bộ";
                dgv_km.Columns[1].Width = 150;
                dgv_km.Columns[2].HeaderText = "Chức vụ";
                dgv_km.Columns[3].HeaderText = "Vai trò";





            }
            catch(Exception ex)
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu đồng ban tổ chức hội thảo {" + ex.Message+"}", "Lỗi");
            }
        }
        private void frm_bomaytochucHTQT_Load(object sender, EventArgs e)
        {
            string ma = Maht;
            loadDLBTC(ma);
            loadDLBTCNT(ma);
            LoadProductListGV();
            LoadProductListKM();
        }

        private void barButtonItem12_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            this.Close();
        }

        private void btn_refresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string ma = Maht;

            txt_tengv.Clear();
            txt_magv.Clear();
            txt_chucvugv.Clear();
            cbo_vaitrogv.SelectedIndex = -1;
            loadDLBTC(ma);

            txt_tenkm.Clear();
            txt_makm.Clear();
            txt_chucvukm.Clear();
            cbo_vaitrokm.SelectedIndex = -1;
            loadDLBTCNT(ma);
           
        }
        public DataTable LayDuLieuBaoCao(string ma)
        {

            string query = " select MaHT,TenHoiThao,NgayToChuc,DiaDiem from HoiThao where CapHoiThao = N'Cấp Quốc Tế' and MaHT = '" +ma+ "' ";
            DataTable dataTable = my.DocDL(query);

            return dataTable;
        }
        public void excelCT()
        {
            try
            {
                string ma = Maht;
                if (string.IsNullOrWhiteSpace(ma))
                {
                    MessageBox.Show("Vui lòng chọn hội thảo ", "Thông báo");
                }
                else
                {
                    DataTable dataTable = LayDuLieuBaoCao(ma);


                    Excel.Application oExcel = new Excel.Application();
                    Excel.Workbook workbook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    Excel.Worksheet oSheet = (Excel.Worksheet)workbook.Worksheets[1];




                    Excel.Range head = oSheet.get_Range("A1", "F1");

                    head.MergeCells = true;

                    head.Value2 = " CHI TIẾT BAN TỔ CHỨC HỘI THẢO QUỐC TẾ ";

                    head.Font.Bold = true;

                    head.Font.Name = "Times New Roman";

                    head.Font.Size = "20";

                    head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    Excel.Range cl1 = oSheet.get_Range("A3", "A3");
                    cl1.Value = "Mã hội thảo";

                    Excel.Range cl2 = oSheet.get_Range("B3", "B3");
                    cl2.Value = "Tên hội thảo";

                    Excel.Range cl3 = oSheet.get_Range("C3", "C3");
                    cl3.Value = "Ngày tổ chức";

                    Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                    cl4.Value = "Địa điểm";

                    Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                    cl5.Value = "Ban tổ chức trong trường";

                    Excel.Range cl10 = oSheet.get_Range("F3", "F3");
                    cl10.Value = "Ban tổ chức ngoài trường";

                    //Excel.Range cl6 = oSheet.get_Range("G3", "G3");
                    //cl6.Value = "Ban lãnh đạo trong trường";

                    //Excel.Range cl7 = oSheet.get_Range("H3", "H3");
                    //cl7.Value = "Ban lãnh đạo ngoài trường";

                    //Excel.Range cl8 = oSheet.get_Range("I3", "I3");
                    //cl8.Value = "Ban tổ chức cuộc thi";

                    //Excel.Range cl9 = oSheet.get_Range("J3", "J3");
                    //cl9.Value = "Ban hỗ trợ kỹ thuật và thư ký";





                    Excel.Range rowHead = oSheet.get_Range("A3", "F3");
                    rowHead.Font.Bold = true;
                    rowHead.Font.Name = "Times New Roman";
                    rowHead.Font.Size = 13;
                    rowHead.Borders.LineStyle = Excel.Constants.xlSolid;
                    rowHead.Interior.ColorIndex = 6;
                    rowHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    // Sau đó, thêm dữ liệu từ DataTable
                    int line = 4;
                    int lines = 4;
                    string maHT;
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {

                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            oSheet.Cells[i + line, j + 1] = dataTable.Rows[i][j];
                            oSheet.Cells[i + line, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            oSheet.Cells[i + line, j + 1].Borders.LineStyle = Excel.Constants.xlSolid;
                            oSheet.Cells[i + line, j + 1].Font.Name = "Times New Roman";

                        }


                        maHT = dataTable.Rows[i][0].ToString();
                        //
                        string query = @" select BanToChucHT.MaGV,GiangVien.HoTen,BanToChucHT.ChucVu,BanToChucHT.VaiTro from BanToChucHT,GiangVien
                                        WHERE BanToChucHT.MaGV = GiangVien.MaGV and BanToChucHT.MaHT = '" + maHT + "'  ";

                        DataTable dt = my.DocDL(query);

                        Excel.Range line1 = oSheet.get_Range("E" + (lines).ToString(), "E" + (lines).ToString());
                        Excel.Range line2 = oSheet.get_Range("F" + (lines).ToString(), "F" + (lines).ToString());
                        //Excel.Range line3 = oSheet.get_Range("I" + (lines).ToString(), "I" + (lines).ToString());
                        //Excel.Range line4 = oSheet.get_Range("J" + (lines).ToString(), "J" + (lines).ToString());

                        for (int row = 0; row < dt.Rows.Count; row++)
                        {
                            string maDT = dt.Rows[row][0].ToString();

                            string cel = dt.Rows[row]["MaGV"].ToString() + "-" + dt.Rows[row]["HoTen"].ToString() + "-" + dt.Rows[row]["ChucVu"].ToString() + "-" + dt.Rows[row]["VaiTro"].ToString() + "\n";
                            line1.Value += cel;

                        }
                        line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line1.Borders.LineStyle = Excel.Constants.xlSolid;
                        line1.Font.Name = "Times New Roman";
                        //
                        //
                        string km = @" select BanToChucHTNT.MaKM,TVNgoaiTruong.HoTen,BanToChucHTNT.ChucVu,BanToChucHTNT.VaiTro from BanToChucHTNT,TVNgoaiTruong
                                        WHERE BanToChucHTNT.MaKM = TVNgoaiTruong.MaKM and BanToChucHTNT.MaHT = '" + ma + "'  ";

                        DataTable dt_ldnt = my.DocDL(km);



                        for (int r = 0; r < dt_ldnt.Rows.Count; r++)
                        {

                            string celSV = dt_ldnt.Rows[r]["MaKM"].ToString() + "-" + dt_ldnt.Rows[r]["HoTen"].ToString() + "-" + dt_ldnt.Rows[r]["ChucVu"].ToString() + "-" + dt_ldnt.Rows[r]["VaiTro"].ToString() + "\n";
                            line2.Value += celSV;


                        }
                        line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line2.Borders.LineStyle = Excel.Constants.xlSolid;
                        line2.Font.Name = "Times New Roman";

                        //
                        //
                        
                        //
                        lines++;




                    }

                    oSheet.Name = "BMTCCT";
                    oExcel.Columns.AutoFit();

                    workbook.Activate();
                    SaveFileDialog saveFile = new SaveFileDialog();
                    if (saveFile.ShowDialog() == DialogResult.OK)
                    {
                        saveFile.Filter = "Text Files|*.xlxs|All Files|*.*";
                        workbook.SaveAs(saveFile.FileName.ToLower());
                        MessageBox.Show("Xuất danh sách thành công", "Thông báo");
                    }

                    oExcel.Quit();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi xuất báo cáo: {ex.Message}");
            }
        }
        private void barButtonItem11_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            excelCT();
        }
        public bool KtraMaTCNT(string magv, string ma)
        {
            try
            {
                string sql = "select * from BanToChucHT where MaGV = '" + magv + "' and MaHT = '" + ma + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }
            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã ban tổ chức !", "Thông báo");
            }
            return true;
        }
        public bool KtraMaTCKM(string makm, string ma)
        {
            try
            {
                string sql = "select * from BanToChucHTNT where MaKM = '" + makm + "' and MaHT = '" + ma + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }
            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã ban tổ chức ngoài trường !", "Thông báo");
            }
            return true;
        }
        private void btn_joinld_Click(object sender, EventArgs e)
        {
            string ma = Maht;
            if(string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng chọn hội thảo cần thêm ban tổ chức", "Thông báo");
            }
            else
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_magv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvugv.Text) || string.IsNullOrWhiteSpace(cbo_vaitrogv.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                    }
                    else
                    {
                        if (KtraMaTCNT(txt_magv.Text, ma))
                        {
                            string sql = "insert into BanToChucHT values (@Magv,@Maht,@Chucvu,@Vaitro) ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Magv", txt_magv.Text);
                            command.Parameters.AddWithValue("@Maht", ma);
                            command.Parameters.AddWithValue("@Chucvu", txt_chucvugv.Text);
                            command.Parameters.AddWithValue("@Vaitro", cbo_vaitrogv.Text);
                            //command.Parameters.AddWithValue("@BanCT", "Ban tổ chức");
                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Thêm ban tổ chúc thành công ", "Thông báo");
                                txt_tengv.Clear();
                                txt_magv.Clear();
                                txt_chucvugv.Clear();
                                cbo_vaitrogv.SelectedIndex = -1;
                                loadDLBTC(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Giảng viên này đã tham gia ban tổ chức", "Thông báo");
                        }
                    }
                    
                }
                catch
                {
                    MessageBox.Show("Lỗi thêm ban tổ chức", "Lỗi");
                }
                
            }
        }

        private void btn_xoald_Click(object sender, EventArgs e)
        {
            string ma = Maht;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng chọn hội thảo cần thêm ban tổ chức", "Thông báo");
            }
            else
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_magv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvugv.Text) || string.IsNullOrWhiteSpace(cbo_vaitrogv.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                    }
                    else
                    {
                        if (!KtraMaTCNT(txt_magv.Text, ma))
                        {
                            string sql = "delete from BanToChucHT where MaGV = @Magv and MaHT=@Maht ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Magv", txt_magv.Text);
                            command.Parameters.AddWithValue("@Maht", ma);
                            //command.Parameters.AddWithValue("@Chucvu", txt_chucvugv.Text);
                            //command.Parameters.AddWithValue("@Vaitro", cbo_vaitrogv.Text);
                            //command.Parameters.AddWithValue("@BanCT", "Ban tổ chức");
                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa ban tổ chúc thành công ", "Thông báo");
                                txt_tengv.Clear();
                                txt_magv.Clear();
                                txt_chucvugv.Clear();
                                cbo_vaitrogv.SelectedIndex = -1;
                                loadDLBTC(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Giảng viên này không tham gia ban tổ chức", "Thông báo");
                        }
                    }
                   
                }
                catch
                {
                    MessageBox.Show("Lỗi xóa ban tổ chức", "Lỗi");
                }

            }
        }

        private void btn_suald_Click(object sender, EventArgs e)
        {
            string ma = Maht;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng chọn hội thảo cần thêm ban tổ chức", "Thông báo");
            }
            else
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_magv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvugv.Text) || string.IsNullOrWhiteSpace(cbo_vaitrogv.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                    }
                    else
                    {
                        if (!KtraMaTCNT(txt_magv.Text, ma))
                        {
                            string sql = "update BanToChucHT set ChucVu=@Chucvu,VaiTro=@Vaitro where MaGV = @Magv and MaHT = @Maht ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Magv", txt_magv.Text);
                            command.Parameters.AddWithValue("@Maht", ma);
                            command.Parameters.AddWithValue("@Chucvu", txt_chucvugv.Text);
                            command.Parameters.AddWithValue("@Vaitro", cbo_vaitrogv.Text);
                            //command.Parameters.AddWithValue("@BanCT", "Ban tổ chức");
                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa ban tổ chúc thành công ", "Thông báo");
                                txt_tengv.Clear();
                                txt_magv.Clear();
                                txt_chucvugv.Clear();
                                cbo_vaitrogv.SelectedIndex = -1;
                                loadDLBTC(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Giảng viên này không tham gia ban tổ chức", "Thông báo");
                        }
                    }
                    
                }
                catch
                {
                    MessageBox.Show("Lỗi sửa ban tổ chức", "Lỗi");
                }

            }
        }

        private void btn_joinldnt_Click(object sender, EventArgs e)
        {
            string ma = Maht;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng chọn hội thảo cần thêm ban tổ chức", "Thông báo");
            }
            else
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_makm.Text) || string.IsNullOrWhiteSpace(txt_tenkm.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvukm.Text) || string.IsNullOrWhiteSpace(cbo_vaitrokm.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                    }
                    else
                    {
                        if (KtraMaTCKM(txt_makm.Text, ma))
                        {
                            string sql = "insert into BanToChucHTNT values (@Makm,@Maht,@Chucvu,@Vaitro) ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Makm", txt_makm.Text);
                            command.Parameters.AddWithValue("@Maht", ma);
                            command.Parameters.AddWithValue("@Chucvu", txt_chucvukm.Text);
                            command.Parameters.AddWithValue("@Vaitro", cbo_vaitrokm.Text);
                            //command.Parameters.AddWithValue("@BanCT", "Ban tổ chức");
                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Thêm ban tổ chúc thành công ", "Thông báo");
                                txt_tenkm.Clear();
                                txt_makm.Clear();
                                txt_chucvukm.Clear();
                                cbo_vaitrokm.SelectedIndex = -1;
                                loadDLBTCNT(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Cán bộ đã tham gia ban tổ chức", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi thêm ban tổ chức", "Lỗi");
                }

            }
        }

        private void btn_xoaldnt_Click(object sender, EventArgs e)
        {
            string ma = Maht;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng chọn hội thảo cần thêm ban tổ chức", "Thông báo");
            }
            else
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_makm.Text) || string.IsNullOrWhiteSpace(txt_tenkm.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvukm.Text) || string.IsNullOrWhiteSpace(cbo_vaitrokm.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                    }
                    else
                    {
                        if (!KtraMaTCKM(txt_makm.Text, ma))
                        {
                            string sql = "delete from BanToChucHTNT where MaKM=@Makm and MaHT=@Maht ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Makm", txt_makm.Text);
                            command.Parameters.AddWithValue("@Maht", ma);
                            //command.Parameters.AddWithValue("@Chucvu", txt_chucvukm.Text);
                            //command.Parameters.AddWithValue("@Vaitro", cbo_vaitrokm.Text);
                            //command.Parameters.AddWithValue("@BanCT", "Ban tổ chức");
                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa ban tổ chúc thành công ", "Thông báo");
                                txt_tenkm.Clear();
                                txt_makm.Clear();
                                txt_chucvukm.Clear();
                                cbo_vaitrokm.SelectedIndex = -1;
                                loadDLBTCNT(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Cán bộ không tham gia ban tổ chức", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi xóa ban tổ chức", "Lỗi");
                }

            }
        }

        private void btn_sualdnt_Click(object sender, EventArgs e)
        {
            string ma = Maht;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng chọn hội thảo cần thêm ban tổ chức", "Thông báo");
            }
            else
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_makm.Text) || string.IsNullOrWhiteSpace(txt_tenkm.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvukm.Text) || string.IsNullOrWhiteSpace(cbo_vaitrokm.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                    }
                    else
                    {
                        if (!KtraMaTCKM(txt_makm.Text, ma))
                        {
                            string sql = "update BanToChucHTNT set ChucVu=@Chucvu,VaiTro=@Vaitro where MaKM= @Makm and MaHT =@Maht ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Makm", txt_makm.Text);
                            command.Parameters.AddWithValue("@Maht", ma);
                            command.Parameters.AddWithValue("@Chucvu", txt_chucvukm.Text);
                            command.Parameters.AddWithValue("@Vaitro", cbo_vaitrokm.Text);
                            //command.Parameters.AddWithValue("@BanCT", "Ban tổ chức");
                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa ban tổ chúc thành công ", "Thông báo");
                                txt_tenkm.Clear();
                                txt_makm.Clear();
                                txt_chucvukm.Clear();
                                cbo_vaitrokm.SelectedIndex = -1;
                                loadDLBTCNT(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Cán bộ không tham gia ban tổ chức", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi sửa ban tổ chức", "Lỗi");
                }

            }
        
        }
        private void LoadProductListGV()
        {
            try
            {

                productListGV = new List<string>();
                string query = "SELECT MaGV FROM GiangVien";
                DataTable tb = my.DocDL(query);
                if (tb.Rows.Count > 0)
                {
                    for (int i = 0; i < tb.Rows.Count; i++)
                    {
                        string ma = tb.Rows[i][0].ToString();
                        productListGV.Add(ma);
                    }
                }
            }
            catch
            {
                MessageBox.Show($"Lỗi thực hiện tạo danh sách giảng viên", "Lỗi");
            }

        }
        private void ShowSuggestionsGV(List<string> suggestions)
        {
            list_nt.Items.Clear();
            list_nt.Items.AddRange(suggestions.ToArray());

            list_nt.Visible = suggestions.Any();
        }
        private void txt_magv_TextChanged(object sender, EventArgs e)
        {
            string searchTerm = txt_magv.Text.ToLower();
            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                List<string> filteredProducts = productListGV
               .Where(product => product.ToLower().Contains(searchTerm))
               .ToList();

                if (filteredProducts != null)
                {
                    ShowSuggestionsGV(filteredProducts);
                }
                else
                {
                    list_nt.Visible = false;

                }


            }
            else
            {
                list_nt.Visible = false;
                txt_tengv.Clear();
            }
        }

        private void list_nt_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_nt.SelectedItem != null)
            {
                string selectedProduct = list_nt.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_magv.Text = selectedProduct;
                    list_nt.Visible = false;
                    string sql = "select HoTen from GiangVien where MaGV = '" + selectedProduct + "' ";
                    DataTable tb = my.DocDL(sql);
                    if (tb.Rows.Count > 0)
                    {
                        string hoten = tb.Rows[0][0].ToString();
                        txt_tengv.Text = hoten;
                    }

                }

            }
        }
        private void LoadProductListKM()
        {
            try
            {

                productListKM = new List<string>();
                string query = "SELECT MaKM FROM TVNgoaiTruong";
                DataTable tb = my.DocDL(query);
                if (tb.Rows.Count > 0)
                {
                    for (int i = 0; i < tb.Rows.Count; i++)
                    {
                        string ma = tb.Rows[i][0].ToString();
                        productListKM.Add(ma);
                    }
                }
            }
            catch
            {
                MessageBox.Show($"Lỗi thực hiện tạo danh sách khách mời", "Lỗi");
            }

        }
        private void ShowSuggestionsKM(List<string> suggestions)
        {
            list_km.Items.Clear();
            list_km.Items.AddRange(suggestions.ToArray());

            list_km.Visible = suggestions.Any();
        }
        private void txt_makm_TextChanged(object sender, EventArgs e)
        {
            string searchTerm = txt_makm.Text.ToLower();
            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                List<string> filteredProducts = productListKM
               .Where(product => product.ToLower().Contains(searchTerm))
               .ToList();

                if (filteredProducts != null)
                {
                    ShowSuggestionsKM(filteredProducts);
                }
                else
                {
                    list_km.Visible = false;

                }


            }
            else
            {
                list_km.Visible = false;
                txt_tenkm.Clear();
            }
        }

        private void list_km_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_km.SelectedItem != null)
            {
                string selectedProduct = list_km.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_makm.Text = selectedProduct;
                    list_km.Visible = false;
                    string sql = "select HoTen from TVNgoaiTruong where MaKM = '" + selectedProduct + "' ";
                    DataTable tb = my.DocDL(sql);
                    if (tb.Rows.Count > 0)
                    {
                        string hoten = tb.Rows[0][0].ToString();
                        txt_tenkm.Text = hoten;
                    }

                }

            }
        }

        private void dgv_nt_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_magv.Text = dgv_nt.CurrentRow.Cells[0].Value.ToString();
                txt_tengv.Text = dgv_nt.CurrentRow.Cells[1].Value.ToString();
                txt_chucvugv.Text = dgv_nt.CurrentRow.Cells[2].Value.ToString();
                cbo_vaitrogv.Text = dgv_nt.CurrentRow.Cells[3].Value.ToString();

            }
            catch
            {
                MessageBox.Show("Lỗi hiển thị dữ liệu ban tổ chức  ", "Lỗi");
            }
        }

        private void dgv_km_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_makm.Text = dgv_km.CurrentRow.Cells[0].Value.ToString();
                txt_tenkm.Text = dgv_km.CurrentRow.Cells[1].Value.ToString();
                txt_chucvukm.Text = dgv_km.CurrentRow.Cells[2].Value.ToString();
                cbo_vaitrokm.Text = dgv_km.CurrentRow.Cells[3].Value.ToString();

            }
            catch
            {
                MessageBox.Show("Lỗi hiển thị dữ liệu ban tổ chức ngoài trường  ", "Lỗi");
            }
        }
    }
}