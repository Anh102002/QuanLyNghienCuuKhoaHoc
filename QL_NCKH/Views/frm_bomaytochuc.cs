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
    public partial class frm_bomaytochuc : DevExpress.XtraEditors.XtraForm
    {
        public frm_bomaytochuc()
        {
            InitializeComponent();
        }

        MyClass my = new MyClass();
        private string mact;
        public string Mact
        {
            get { return this.mact; }
            set { this.mact = value; }
        }
        public void loadDLBLD(string ma)
        {
            try
            {

                string query = @" select BanLanhDaoCT.MaGV,GiangVien.HoTen,BanLanhDaoCT.ChucVu,BanLanhDaoCT.VaiTro from BanLanhDaoCT,GiangVien
                                        WHERE BanLanhDaoCT.MaGV = GiangVien.MaGV and BanLanhDaoCT.MaCuocThi = '"+ma+ "'  ";
                DataTable dt = my.DocDL(query);
                dgv_ld.DataSource = dt;
                dgv_ld.Columns[0].HeaderText = "Mã giảng viên";
                dgv_ld.Columns[1].HeaderText = "Tên giảng viên";
                dgv_ld.Columns[1].Width = 150;
                dgv_ld.Columns[2].HeaderText = "Chức vụ";
                dgv_ld.Columns[3].HeaderText = "Vai trò";
               




            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu ban lãnh đạo cuộc thi {1}", "Lỗi");
            }
        }

        public void loadDLBLDNT(string ma)
        {
            try
            {

                string query = @" select BanLanhDaoCTNT.MaKM,TVNgoaiTruong.HoTen,BanLanhDaoCTNT.ChucVu,BanLanhDaoCTNT.VaiTro from BanLanhDaoCTNT,TVNgoaiTruong
                                        WHERE BanLanhDaoCTNT.MaKM = TVNgoaiTruong.MaKM and BanLanhDaoCTNT.MaCuocThi = '" + ma + "' ";
                DataTable dt = my.DocDL(query);
                dgv_ldnt.DataSource = dt;
                dgv_ldnt.Columns[0].HeaderText = "Mã thành viên";
                dgv_ldnt.Columns[1].HeaderText = "Tên thành viên";
                dgv_ldnt.Columns[1].Width = 150;
                dgv_ldnt.Columns[2].HeaderText = "Chức vụ";
                dgv_ldnt.Columns[3].HeaderText = "Vai trò";





            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu ban lãnh đạo cuộc thi {2}", "Lỗi");
            }
        }
        public void loadDLBTC(string ma)
        {
            try
            {

                string query = @" select BanToChucCT.MaGV,GiangVien.HoTen,BanToChucCT.ChucVu,BanToChucCT.VaiTro from BanToChucCT,GiangVien
                                        WHERE BanToChucCT.MaGV = GiangVien.MaGV and BanToChucCT.MaCuocThi = '" + ma + "'  ";
                DataTable dt = my.DocDL(query);
                dgv_tc.DataSource = dt;
                dgv_tc.Columns[0].HeaderText = "Mã giảng viên";
                dgv_tc.Columns[1].HeaderText = "Tên giảng viên";
                dgv_tc.Columns[1].Width = 150;
                dgv_tc.Columns[2].HeaderText = "Chức vụ";
                dgv_tc.Columns[3].HeaderText = "Vai trò";





            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu ban tổ chức cuộc thi ", "Lỗi");
            }
        }

        public void loadDLBHT(string ma)
        {
            try
            {

                string query = @" select BanHoTroCT.MaGV,GiangVien.HoTen,BanHoTroCT.ChucVu,BanHoTroCT.VaiTro from BanHoTroCT,GiangVien
                                        WHERE BanHoTroCT.MaGV = GiangVien.MaGV and BanHoTroCT.MaCuocThi = '" + ma + "'  ";
                DataTable dt = my.DocDL(query);
                dgv_kt.DataSource = dt;
                dgv_kt.Columns[0].HeaderText = "Mã giảng viên";
                dgv_kt.Columns[1].HeaderText = "Tên giảng viên";
                dgv_kt.Columns[1].Width = 150;
                dgv_kt.Columns[2].HeaderText = "Chức vụ";
                dgv_kt.Columns[3].HeaderText = "Vai trò";





            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu ban hỗ trợ kỹ thuật cuộc thi ", "Lỗi");
            }
        }
        private void frm_bomaytochuc_Load(object sender, EventArgs e)
        {
            string ma = Mact;
            loadDLBLD(ma);
            loadDLBLDNT(ma);
            loadDLBTC(ma);
            loadDLBHT(ma);
            LoadProductListGV();
            LoadProductListKM();

        }
        public bool KtraMaLD(string magv,string ma)
        {
            try
            {
                string sql = "select * from BanLanhDaoCT where MaGV = '" + magv + "' and MaCuocThi = '" + ma + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }
            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã ban lãnh đạo !", "Thông báo");
            }
            return true;
        }
        private void btn_joinld_Click(object sender, EventArgs e)
        {
            string ma = Mact;
            if(string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng quay lại chọn cuộc thi ", "Thông báo");
            }
            else
            {
                if(string.IsNullOrWhiteSpace(txt_mald.Text) || string.IsNullOrWhiteSpace(txt_tenld.Text) 
                    || string.IsNullOrWhiteSpace(txt_chucvuld.Text) || string.IsNullOrWhiteSpace(cbo_vaitrold.Text)  )
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                }
                else
                {
                    try
                    {
                        
                            
                                    if (KtraMaLD(txt_mald.Text, ma))
                                    {
                                        string sql = "insert into BanLanhDaoCT values (@Magv,@Mact,@Chucvu,@Vaitro) ";
                                        SqlCommand command = my.SqlCommand(sql);
                                        command.Parameters.AddWithValue("@Magv", txt_mald.Text);
                                        command.Parameters.AddWithValue("@Mact", ma);
                                        command.Parameters.AddWithValue("@Chucvu", txt_chucvuld.Text);
                                        command.Parameters.AddWithValue("@Vaitro", cbo_vaitrold.Text);
                                        //command.Parameters.AddWithValue("@BanCT", "Ban lãnh đạo");
                                        int up = command.ExecuteNonQuery();
                                        if (up > 0)
                                        {
                                            MessageBox.Show("Thêm ban lãnh đạo thành công ", "Thông báo");
                                            txt_tenld.Clear();
                                            txt_mald.Clear();
                                            txt_chucvuld.Clear();
                                            cbo_vaitrold.SelectedIndex = -1;
                                            loadDLBLD(ma);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Đã có trong ban lãnh đạo này", "Thông báo");
                                    }
                                
                           


                    }
                    catch
                    {
                        MessageBox.Show("Lỗi thêm ban lãnh đạo ", "Lỗi");
                    }
                }
            }
        }

        private void barButtonItem12_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            this.Close();

        }

        private void btn_xoald_Click(object sender, EventArgs e)
        {
            string ma = Mact;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng quay lại chọn cuộc thi ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(txt_mald.Text) || string.IsNullOrWhiteSpace(txt_tenld.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvuld.Text) || string.IsNullOrWhiteSpace(cbo_vaitrold.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                }
                else
                {
                    try
                    {
                        if (!KtraMaLD(txt_mald.Text,ma))
                        {
                            string sql = "delete from BanLanhDaoCT where MaGV=@Magv and MaCuocThi = @Mact ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Magv", txt_mald.Text);
                            command.Parameters.AddWithValue("@Mact", ma);
                            //command.Parameters.AddWithValue("@Chucvu", txt_chucvuld.Text);
                            //command.Parameters.AddWithValue("@Vaitro", cbo_vaitrold.Text);
                            //command.Parameters.AddWithValue("@BanCT", "Ban lãnh đạo");
                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa ban lãnh đạo thành công ", "Thông báo");
                                txt_tenld.Clear();
                                txt_mald.Clear();
                                txt_chucvuld.Clear();
                                cbo_vaitrold.SelectedIndex = -1;
                                loadDLBLD(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("không có trong ban lãnh đạo này", "Thông báo");
                        }

                    }
                    catch
                    {
                        MessageBox.Show("Lỗi xóa ban lãnh đạo ", "Lỗi");
                    }
                }
            }
        }

        private void btn_suald_Click(object sender, EventArgs e)
        {
            string ma = Mact;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng quay lại chọn cuộc thi ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(txt_mald.Text) || string.IsNullOrWhiteSpace(txt_tenld.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvuld.Text) || string.IsNullOrWhiteSpace(cbo_vaitrold.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                }
                else
                {
                    try
                    {
                        if (!KtraMaLD(txt_mald.Text,ma))
                        {
                            string sql = "update BanLanhDaoCT set ChucVu = @Chucvu,VaiTro=@Vaitro  where MaGV=@Magv and MaCuocThi = @Mact ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Magv", txt_mald.Text);
                            command.Parameters.AddWithValue("@Mact", ma);
                            command.Parameters.AddWithValue("@Chucvu", txt_chucvuld.Text);
                            command.Parameters.AddWithValue("@Vaitro", cbo_vaitrold.Text);
                            //command.Parameters.AddWithValue("@BanCT", "Ban lãnh đạo");
                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa ban lãnh đạo thành công ", "Thông báo");
                                txt_tenld.Clear();
                                txt_mald.Clear();
                                txt_chucvuld.Clear();
                                cbo_vaitrold.SelectedIndex = -1;
                                loadDLBLD(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("không có trong ban lãnh đạo này", "Thông báo");
                        }

                    }
                    catch
                    {
                        MessageBox.Show("Lỗi sửa ban lãnh đạo ", "Lỗi");
                    }
                }
            }
        }

        private void dgv_ld_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_mald.Text = dgv_ld.CurrentRow.Cells[0].Value.ToString();
                txt_tenld.Text = dgv_ld.CurrentRow.Cells[1].Value.ToString();
                txt_chucvuld.Text = dgv_ld.CurrentRow.Cells[2].Value.ToString();
                cbo_vaitrold.Text = dgv_ld.CurrentRow.Cells[3].Value.ToString();
               
            }
            catch
            {
                MessageBox.Show("Lỗi hiển thị dữ liệu ban lãnh đạo ", "Lỗi");
            }
        }

        private void dgv_ldnt_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_maldnt.Text = dgv_ldnt.CurrentRow.Cells[0].Value.ToString();
                txt_tenldnt.Text = dgv_ldnt.CurrentRow.Cells[1].Value.ToString();
                txt_chucvunt.Text = dgv_ldnt.CurrentRow.Cells[2].Value.ToString();
                cbo_vaitroldnt.Text = dgv_ldnt.CurrentRow.Cells[3].Value.ToString();

            }
            catch
            {
                MessageBox.Show("Lỗi hiển thị dữ liệu ban lãnh đạo ngoài trường", "Lỗi");
            }
        }

        private void dgv_tc_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_mabtc.Text = dgv_tc.CurrentRow.Cells[0].Value.ToString();
                txt_tenbtc.Text = dgv_tc.CurrentRow.Cells[1].Value.ToString();
                txt_chucvubtc.Text = dgv_tc.CurrentRow.Cells[2].Value.ToString();
                cbo_vaitrobtc.Text = dgv_tc.CurrentRow.Cells[3].Value.ToString();

            }
            catch
            {
                MessageBox.Show("Lỗi hiển thị dữ liệu ban tổ chức", "Lỗi");
            }
        }

        private void dgv_kt_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_mabkt.Text = dgv_kt.CurrentRow.Cells[0].Value.ToString();
                txt_tenbkt.Text = dgv_kt.CurrentRow.Cells[1].Value.ToString();
                txt_chucvubkt.Text = dgv_kt.CurrentRow.Cells[2].Value.ToString();
                cbo_vaitrobkt.Text = dgv_kt.CurrentRow.Cells[3].Value.ToString();

            }
            catch
            {
                MessageBox.Show("Lỗi hiển thị dữ liệu ban hỗ trợ", "Lỗi");
            }
        }
        public bool KtraMaLDNT(string makm,string ma)
        {
            try
            {
                string sql = "select * from BanLanhDaoCTNT where MaKM = '" + makm + "' and MaCuocThi = '" + ma + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }
            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã ban lãnh đạo ngoài trường!", "Thông báo");
            }
            return true;
        }
        private void btn_joinldnt_Click(object sender, EventArgs e)
        {
            string ma = Mact;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng quay lại chọn cuộc thi ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(txt_maldnt.Text) || string.IsNullOrWhiteSpace(txt_tenldnt.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvunt.Text) || string.IsNullOrWhiteSpace(cbo_vaitroldnt.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                }
                else
                {
                    try
                    {
                        if (KtraMaLDNT(txt_maldnt.Text,ma))
                        {
                            string sql = "insert into BanLanhDaoCTNT values (@Makm,@Mact,@Chucvu,@Vaitro) ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Makm", txt_maldnt.Text);
                            command.Parameters.AddWithValue("@Mact", ma);
                            command.Parameters.AddWithValue("@Chucvu", txt_chucvunt.Text);
                            command.Parameters.AddWithValue("@Vaitro", cbo_vaitroldnt.Text);
                            
                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Thêm ban lãnh đạo thành công ", "Thông báo");
                                txt_tenldnt.Clear();
                                txt_maldnt.Clear();
                                txt_chucvunt.Clear();
                                cbo_vaitroldnt.SelectedIndex = -1;
                                loadDLBLDNT(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Đã có trong ban lãnh đạo này", "Thông báo");
                        }

                    }
                    catch
                    {
                        MessageBox.Show("Lỗi thêm ban lãnh đạo ", "Lỗi");
                    }
                }
            }
        }

        private void btn_xoaldnt_Click(object sender, EventArgs e)
        {
            string ma = Mact;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng quay lại chọn cuộc thi ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(txt_maldnt.Text) || string.IsNullOrWhiteSpace(txt_tenldnt.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvunt.Text) || string.IsNullOrWhiteSpace(cbo_vaitroldnt.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                }
                else
                {
                    try
                    {
                        if (!KtraMaLDNT(txt_maldnt.Text,ma))
                        {
                            string sql = "delete from BanLanhDaoCTNT where  MaKM = @Makm and MaCuocThi = @Mact  ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Makm", txt_maldnt.Text);
                            command.Parameters.AddWithValue("@Mact", ma);
                            //command.Parameters.AddWithValue("@Chucvu", txt_chucvunt.Text);
                            //command.Parameters.AddWithValue("@Vaitro", cbo_vaitroldnt.Text);

                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa ban lãnh đạo thành công ", "Thông báo");
                                txt_tenldnt.Clear();
                                txt_maldnt.Clear();
                                txt_chucvunt.Clear();
                                cbo_vaitroldnt.SelectedIndex = -1;
                                loadDLBLDNT(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("không có trong ban lãnh đạo này", "Thông báo");
                        }

                    }
                    catch
                    {
                        MessageBox.Show("Lỗi xóa ban lãnh đạo ", "Lỗi");
                    }
                }
            }
        }

        private void btn_sualdnt_Click(object sender, EventArgs e)
        {
            string ma = Mact;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng quay lại chọn cuộc thi ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(txt_maldnt.Text) || string.IsNullOrWhiteSpace(txt_tenldnt.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvunt.Text) || string.IsNullOrWhiteSpace(cbo_vaitroldnt.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                }
                else
                {
                    try
                    {
                        if (!KtraMaLDNT(txt_maldnt.Text,ma))
                        {
                            string sql = "update BanLanhDaoCTNT set ChucVu=@Chucvu , VaiTro=@Vaitro where  MaKM = @Makm and MaCuocThi = @Mact  ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Makm", txt_maldnt.Text);
                            command.Parameters.AddWithValue("@Mact", ma);
                            command.Parameters.AddWithValue("@Chucvu", txt_chucvunt.Text);
                            command.Parameters.AddWithValue("@Vaitro", cbo_vaitroldnt.Text);

                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa ban lãnh đạo thành công ", "Thông báo");
                                txt_tenldnt.Clear();
                                txt_maldnt.Clear();
                                txt_chucvunt.Clear();
                                cbo_vaitroldnt.SelectedIndex = -1;
                                loadDLBLDNT(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("không có trong ban lãnh đạo này", "Thông báo");
                        }

                    }
                    catch
                    {
                        MessageBox.Show("Lỗi sửa ban lãnh đạo ", "Lỗi");
                    }
                }
            }
        }
        public bool KtraMaTC(string magv,string ma)
        {
            try
            {
                string sql = "select * from BanToChucCT where MaGV = '" + magv+ "' and MaCuocThi = '" + ma + "' ";
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
        private void btn_jointc_Click(object sender, EventArgs e)
        {
            string ma = Mact;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng quay lại chọn cuộc thi ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(txt_mabtc.Text) || string.IsNullOrWhiteSpace(txt_tenbtc.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvubtc.Text) || string.IsNullOrWhiteSpace(cbo_vaitrobtc.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                }
                else
                {
                    try
                    {
                        
                                if (KtraMaTC(txt_mabtc.Text, ma))
                                {
                                    string sql = "insert into BanToChucCT values (@Magv,@Mact,@Chucvu,@Vaitro) ";
                                    SqlCommand command = my.SqlCommand(sql);
                                    command.Parameters.AddWithValue("@Magv", txt_mabtc.Text);
                                    command.Parameters.AddWithValue("@Mact", ma);
                                    command.Parameters.AddWithValue("@Chucvu", txt_chucvubtc.Text);
                                    command.Parameters.AddWithValue("@Vaitro", cbo_vaitrobtc.Text);
                                    //command.Parameters.AddWithValue("@BanCT", "Ban tổ chức");
                                    int up = command.ExecuteNonQuery();
                                    if (up > 0)
                                    {
                                        MessageBox.Show("Thêm ban tổ chúc thành công ", "Thông báo");
                                        txt_tenbtc.Clear();
                                        txt_mabtc.Clear();
                                        txt_chucvubtc.Clear();
                                        cbo_vaitrobtc.SelectedIndex = -1;
                                        loadDLBTC(ma);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Đã có trong ban tổ chức này", "Thông báo");
                                }
                            


                    }
                    catch
                    {
                        MessageBox.Show("Lỗi thêm ban tổ chức ", "Lỗi");
                    }
                }
            }
        }

        private void btn_xoatc_Click(object sender, EventArgs e)
        {
            string ma = Mact;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng quay lại chọn cuộc thi ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(txt_mabtc.Text) || string.IsNullOrWhiteSpace(txt_tenbtc.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvubtc.Text) || string.IsNullOrWhiteSpace(cbo_vaitrobtc.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                }
                else
                {
                    try
                    {
                        if (!KtraMaTC(txt_mabtc.Text,ma))
                        {
                            string sql = "delete from BanToChucCT where MaGV=@Magv and MaCuocThi=@Mact ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Magv", txt_mabtc.Text);
                            command.Parameters.AddWithValue("@Mact", ma);
                            //command.Parameters.AddWithValue("@Chucvu", txt_chucvubtc.Text);
                            //command.Parameters.AddWithValue("@Vaitro", cbo_vaitrobtc.Text);
                            //command.Parameters.AddWithValue("@BanCT", "Ban tổ chức");
                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa ban tổ chúc thành công ", "Thông báo");
                                txt_tenbtc.Clear();
                                txt_mabtc.Clear();
                                txt_chucvubtc.Clear();
                                cbo_vaitrobtc.SelectedIndex = -1;
                                loadDLBTC(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("không có trong ban tổ chức này", "Thông báo");
                        }

                    }
                    catch
                    {
                        MessageBox.Show("Lỗi xóa ban tổ chức ", "Lỗi");
                    }
                }
            }
        }

        private void btn_suatc_Click(object sender, EventArgs e)
        {
            string ma = Mact;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng quay lại chọn cuộc thi ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(txt_mabtc.Text) || string.IsNullOrWhiteSpace(txt_tenbtc.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvubtc.Text) || string.IsNullOrWhiteSpace(cbo_vaitrobtc.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                }
                else
                {
                    try
                    {
                        if (!KtraMaTC(txt_mabtc.Text,ma))
                        {
                            string sql = "update BanToChucCT set ChucVu=@Chucvu,VaiTro=@Vaitro where MaGV=@Magv and MaCuocThi=@Mact ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Magv", txt_mabtc.Text);
                            command.Parameters.AddWithValue("@Mact", ma);
                            command.Parameters.AddWithValue("@Chucvu", txt_chucvubtc.Text);
                            command.Parameters.AddWithValue("@Vaitro", cbo_vaitrobtc.Text);
                            //command.Parameters.AddWithValue("@BanCT", "Ban tổ chức");
                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa ban tổ chúc thành công ", "Thông báo");
                                txt_tenbtc.Clear();
                                txt_mabtc.Clear();
                                txt_chucvubtc.Clear();
                                cbo_vaitrobtc.SelectedIndex = -1;
                                loadDLBTC(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("không có trong ban tổ chức này", "Thông báo");
                        }

                    }
                    catch
                    {
                        MessageBox.Show("Lỗi sửa ban tổ chức ", "Lỗi");
                    }
                }
            }
        }
        public bool KtraMaHT(string magv,string ma)
        {
            try
            {
                string sql = "select * from BanHoTroCT where MaGV = '" + magv + "' and MaCuocThi = '" + ma + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }
            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã ban hỗ trợ kỹ thuật !", "Thông báo");
            }
            return true;
        }
        private void btn_joinkt_Click(object sender, EventArgs e)
        {
            string ma = Mact;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng quay lại chọn cuộc thi ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(txt_mabkt.Text) || string.IsNullOrWhiteSpace(txt_tenbkt.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvubkt.Text) || string.IsNullOrWhiteSpace(cbo_vaitrobkt.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                }
                else
                {
                    try
                    {
                        
                                if (KtraMaHT(txt_mabkt.Text, ma))
                                {
                                    string sql = "insert into BanHoTroCT values (@Magv,@Mact,@Chucvu,@Vaitro) ";
                                    SqlCommand command = my.SqlCommand(sql);
                                    command.Parameters.AddWithValue("@Magv", txt_mabkt.Text);
                                    command.Parameters.AddWithValue("@Mact", ma);
                                    command.Parameters.AddWithValue("@Chucvu", txt_chucvubkt.Text);
                                    command.Parameters.AddWithValue("@Vaitro", cbo_vaitrobkt.Text);
                                    //command.Parameters.AddWithValue("@BanCT", "Ban tổ chức");
                                    int up = command.ExecuteNonQuery();
                                    if (up > 0)
                                    {
                                        MessageBox.Show("Thêm ban hỗ trợ kỹ thuật thành công ", "Thông báo");
                                        txt_tenbkt.Clear();
                                        txt_mabkt.Clear();
                                        txt_chucvubkt.Clear();
                                        cbo_vaitrobkt.SelectedIndex = -1;
                                        loadDLBHT(ma);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Đã có trong ban hỗ trợ kỹ thuật này", "Thông báo");
                                }
                           


                    }
                    catch
                    {
                        MessageBox.Show("Lỗi thêm ban hỗ trợ kỹ thuật ", "Lỗi");
                    }
                }
            }
        }

        private void btn_xoakt_Click(object sender, EventArgs e)
        {
            string ma = Mact;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng quay lại chọn cuộc thi ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(txt_mabkt.Text) || string.IsNullOrWhiteSpace(txt_tenbkt.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvubkt.Text) || string.IsNullOrWhiteSpace(cbo_vaitrobkt.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                }
                else
                {
                    try
                    {
                        if (!KtraMaHT(txt_mabkt.Text,ma))
                        {
                            string sql = "delete from BanHoTroCT where  MaGV=@Magv and MaCuocThi=@Mact ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Magv", txt_mabkt.Text);
                            command.Parameters.AddWithValue("@Mact", ma);
                            //command.Parameters.AddWithValue("@Chucvu", txt_chucvubkt.Text);
                            //command.Parameters.AddWithValue("@Vaitro", cbo_vaitrobkt.Text);
                            //command.Parameters.AddWithValue("@BanCT", "Ban tổ chức");
                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa ban hỗ trợ kỹ thuật thành công ", "Thông báo");
                                txt_tenbkt.Clear();
                                txt_mabkt.Clear();
                                txt_chucvubkt.Clear();
                                cbo_vaitrobkt.SelectedIndex = -1;
                                loadDLBHT(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("không có trong ban hỗ trợ kỹ thuật này", "Thông báo");
                        }

                    }
                    catch
                    {
                        MessageBox.Show("Lỗi xóa ban hỗ trợ kỹ thuật ", "Lỗi");
                    }
                }
            }
        }

        private void btn_suakt_Click(object sender, EventArgs e)
        {
            string ma = Mact;
            if (string.IsNullOrWhiteSpace(ma))
            {
                MessageBox.Show("Vui lòng quay lại chọn cuộc thi ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(txt_mabkt.Text) || string.IsNullOrWhiteSpace(txt_tenbkt.Text)
                    || string.IsNullOrWhiteSpace(txt_chucvubkt.Text) || string.IsNullOrWhiteSpace(cbo_vaitrobkt.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                }
                else
                {
                    try
                    {
                        if (!KtraMaHT(txt_mabkt.Text,ma))
                        {
                            string sql = "update BanHoTroCT set ChucVu=@Chucvu,VaiTro=@Vaitro  where  MaGV=@Magv and MaCuocThi=@Mact ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Magv", txt_mabkt.Text);
                            command.Parameters.AddWithValue("@Mact", ma);
                            command.Parameters.AddWithValue("@Chucvu", txt_chucvubkt.Text);
                            command.Parameters.AddWithValue("@Vaitro", cbo_vaitrobkt.Text);
                            
                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa ban hỗ trợ kỹ thuật thành công ", "Thông báo");
                                txt_tenbkt.Clear();
                                txt_mabkt.Clear();
                                txt_chucvubkt.Clear();
                                cbo_vaitrobkt.SelectedIndex = -1;
                                loadDLBHT(ma);
                            }
                        }
                        else
                        {
                            MessageBox.Show("không có trong ban hỗ trợ kỹ thuật này", "Thông báo");
                        }

                    }
                    catch
                    {
                        MessageBox.Show("Lỗi sửa ban hỗ trợ kỹ thuật ", "Lỗi");
                    }
                }
            }
        }

        private void btn_refresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string ma = Mact;

            txt_tenbkt.Clear();
            txt_mabkt.Clear();
            txt_chucvubkt.Clear();
            cbo_vaitrobkt.SelectedIndex = -1;
            loadDLBHT(ma);

            txt_tenbtc.Clear();
            txt_mabtc.Clear();
            txt_chucvubtc.Clear();
            cbo_vaitrobtc.SelectedIndex = -1;
            loadDLBTC(ma);

            txt_tenldnt.Clear();
            txt_maldnt.Clear();
            txt_chucvunt.Clear();
            cbo_vaitroldnt.SelectedIndex = -1;
            loadDLBLDNT(ma);

            txt_tenld.Clear();
            txt_mald.Clear();
            txt_chucvuld.Clear();
            cbo_vaitrold.SelectedIndex = -1;
            loadDLBLD(ma);
        }
        public DataTable LayDuLieuBaoCao(string ma)
        {

            string query = " select MaCuocThi,TenCuocThi,LinhVuc,NamTC,KinhPhi,LoaiCuocThi from CuocThiSTKN where MaCuocThi = '"+ma+"' ";
            DataTable dataTable = my.DocDL(query);

            return dataTable;
        }
        public void excelCT()
        {
            try
            {
                string ma = Mact;
                if(string.IsNullOrWhiteSpace(ma))
                {
                    MessageBox.Show("Vui lòng chọn cuộc thi ", "Thông báo");
                }
                else
                {
                    DataTable dataTable = LayDuLieuBaoCao(ma);


                    Excel.Application oExcel = new Excel.Application();
                    Excel.Workbook workbook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    Excel.Worksheet oSheet = (Excel.Worksheet)workbook.Worksheets[1];




                    Excel.Range head = oSheet.get_Range("A1", "J1");

                    head.MergeCells = true;

                    head.Value2 = " CHI TIẾT BỘ MÁY TỔ CHỨC CUỘC THI   ";

                    head.Font.Bold = true;

                    head.Font.Name = "Times New Roman";

                    head.Font.Size = "20";

                    head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    Excel.Range cl1 = oSheet.get_Range("A3", "A3");
                    cl1.Value = "Mã cuộc thi";

                    Excel.Range cl2 = oSheet.get_Range("B3", "B3");
                    cl2.Value = "Tên cuộc thi";

                    Excel.Range cl3 = oSheet.get_Range("C3", "C3");
                    cl3.Value = "Lĩnh vực";

                    Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                    cl4.Value = "Năm tổ chức";

                    Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                    cl5.Value = "Kinh phí";

                    Excel.Range cl10 = oSheet.get_Range("F3", "F3");
                    cl10.Value = "Loại cuộc thi";

                    Excel.Range cl6 = oSheet.get_Range("G3", "G3");
                    cl6.Value = "Ban lãnh đạo trong trường";

                    Excel.Range cl7 = oSheet.get_Range("H3", "H3");
                    cl7.Value = "Ban lãnh đạo ngoài trường";

                    Excel.Range cl8 = oSheet.get_Range("I3", "I3");
                    cl8.Value = "Ban tổ chức cuộc thi";

                    Excel.Range cl9 = oSheet.get_Range("J3", "J3");
                    cl9.Value = "Ban hỗ trợ kỹ thuật và thư ký";





                    Excel.Range rowHead = oSheet.get_Range("A3", "J3");
                    rowHead.Font.Bold = true;
                    rowHead.Font.Name = "Times New Roman";
                    rowHead.Font.Size = 13;
                    rowHead.Borders.LineStyle = Excel.Constants.xlSolid;
                    rowHead.Interior.ColorIndex = 6;
                    rowHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    // Sau đó, thêm dữ liệu từ DataTable
                    int line = 4;
                    int lines = 4;
                    string maCT;
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {

                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            oSheet.Cells[i + line, j + 1] = dataTable.Rows[i][j];
                            oSheet.Cells[i + line, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            oSheet.Cells[i + line, j + 1].Borders.LineStyle = Excel.Constants.xlSolid;
                            oSheet.Cells[i + line, j + 1].Font.Name = "Times New Roman";

                        }


                        maCT = dataTable.Rows[i][0].ToString();
                        //
                        string query = @" select BanLanhDaoCT.MaGV,GiangVien.HoTen,BanLanhDaoCT.ChucVu,BanLanhDaoCT.VaiTro from BanLanhDaoCT,GiangVien
                                        WHERE BanLanhDaoCT.MaGV = GiangVien.MaGV and BanLanhDaoCT.MaCuocThi = '" + ma + "'  ";

                        DataTable dt = my.DocDL(query);

                        Excel.Range line1 = oSheet.get_Range("G" + (lines).ToString(), "G" + (lines).ToString());
                        Excel.Range line2 = oSheet.get_Range("H" + (lines).ToString(), "H" + (lines).ToString());
                        Excel.Range line3 = oSheet.get_Range("I" + (lines).ToString(), "I" + (lines).ToString());
                        Excel.Range line4 = oSheet.get_Range("J" + (lines).ToString(), "J" + (lines).ToString());

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
                        string ldnt = @" select BanLanhDaoCTNT.MaKM,TVNgoaiTruong.HoTen,BanLanhDaoCTNT.ChucVu,BanLanhDaoCTNT.VaiTro from BanLanhDaoCTNT,TVNgoaiTruong
                                        WHERE BanLanhDaoCTNT.MaKM = TVNgoaiTruong.MaKM and BanLanhDaoCTNT.MaCuocThi = '" + ma + "' ";

                        DataTable dt_ldnt = my.DocDL(ldnt);



                        for (int r = 0; r < dt_ldnt.Rows.Count; r++)
                        {

                            string celSV = dt_ldnt.Rows[r]["MaKM"].ToString() + "-" + dt_ldnt.Rows[r]["HoTen"].ToString() + "-" + dt_ldnt.Rows[r]["ChucVu"].ToString() +"-"+ dt_ldnt.Rows[r]["VaiTro"].ToString() + "\n";
                            line2.Value += celSV;


                        }
                        line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line2.Borders.LineStyle = Excel.Constants.xlSolid;
                        line2.Font.Name = "Times New Roman";

                        //
                        //
                        string btc = @" select BanToChucCT.MaGV,GiangVien.HoTen,BanToChucCT.ChucVu,BanToChucCT.VaiTro from BanToChucCT,GiangVien
                                        WHERE BanToChucCT.MaGV = GiangVien.MaGV and BanToChucCT.MaCuocThi = '" + ma + "'  ";

                        DataTable dt_btc = my.DocDL(btc);



                        for (int r1 = 0; r1 < dt_btc.Rows.Count; r1++)
                        {

                            string celSVNT = dt_btc.Rows[r1]["MaGV"].ToString() + "-" + dt_btc.Rows[r1]["HoTen"].ToString() + "-" + dt_btc.Rows[r1]["ChucVu"].ToString() +"-" + dt_btc.Rows[r1]["VaiTro"].ToString() + "\n";
                            line3.Value += celSVNT;


                        }
                        line3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line3.Borders.LineStyle = Excel.Constants.xlSolid;

                        line3.Font.Name = "Times New Roman";

                        //
                        //
                        string bht = @" select BanHoTroCT.MaGV,GiangVien.HoTen,BanHoTroCT.ChucVu,BanHoTroCT.VaiTro from BanHoTroCT,GiangVien
                                        WHERE BanHoTroCT.MaGV = GiangVien.MaGV and BanHoTroCT.MaCuocThi = '" + ma + "'  ";
                        DataTable dt_bht = my.DocDL(bht);



                        for (int gv = 0; gv < dt_bht.Rows.Count; gv++)
                        {

                            string celGV = dt_bht.Rows[gv]["MaGV"].ToString() + "-" + dt_bht.Rows[gv]["HoTen"].ToString() +"-"+ dt_bht.Rows[gv]["ChucVu"].ToString() +"-"+ dt_bht.Rows[gv]["VaiTro"].ToString() + "\n";
                            line4.Value += celGV;


                        }
                        line4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line4.Borders.LineStyle = Excel.Constants.xlSolid;
                        line4.Font.Name = "Times New Roman";
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
        private List<string> productListGV;
        private List<string> productListKM;

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
            list_ld.Items.Clear();
            list_ld.Items.AddRange(suggestions.ToArray());

            list_ld.Visible = suggestions.Any();
        }
        private void txt_mald_TextChanged(object sender, EventArgs e)
        {
            string searchTerm = txt_mald.Text.ToLower();
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
                    list_ld.Visible = false;

                }


            }
            else
            {
                list_ld.Visible = false;
                txt_tenld.Clear();
            }
        }

        private void list_ld_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_ld.SelectedItem != null)
            {
                string selectedProduct = list_ld.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_mald.Text = selectedProduct;
                    list_ld.Visible = false;
                    string sql = "select HoTen from GiangVien where MaGV = '" + selectedProduct + "' ";
                    DataTable tb = my.DocDL(sql);
                    if (tb.Rows.Count > 0)
                    {
                        string hoten = tb.Rows[0][0].ToString();
                        txt_tenld.Text = hoten;
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
            list_ldnt.Items.Clear();
            list_ldnt.Items.AddRange(suggestions.ToArray());

            list_ldnt.Visible = suggestions.Any();
        }
        private void txt_maldnt_TextChanged(object sender, EventArgs e)
        {
            string searchTerm = txt_maldnt.Text.ToLower();
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
                    list_ldnt.Visible = false;

                }


            }
            else
            {
                list_ldnt.Visible = false;
                txt_tenldnt.Clear();
            }
        }

        private void list_ldnt_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(list_ldnt.SelectedItem != null)
            {
                string selectedProduct = list_ldnt.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_maldnt.Text = selectedProduct;
                    list_ldnt.Visible = false;
                    string sql = "select HoTen from TVNgoaiTruong where MaKM = '" + selectedProduct + "' ";
                    DataTable tb = my.DocDL(sql);
                    if (tb.Rows.Count > 0)
                    {
                        string hoten = tb.Rows[0][0].ToString();
                        txt_tenldnt.Text = hoten;
                    }

                }

            }
        }
        private void ShowSuggestionsTC(List<string> suggestions)
        {
            list_tc.Items.Clear();
            list_tc.Items.AddRange(suggestions.ToArray());

            list_tc.Visible = suggestions.Any();
        }
        private void txt_mabtc_TextChanged(object sender, EventArgs e)
        {
            string searchTerm = txt_mabtc.Text.ToLower();
            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                List<string> filteredProducts = productListGV
               .Where(product => product.ToLower().Contains(searchTerm))
               .ToList();

                if (filteredProducts != null)
                {
                    ShowSuggestionsTC(filteredProducts);
                }
                else
                {
                    list_tc.Visible = false;

                }


            }
            else
            {
                list_tc.Visible = false;
                txt_tenbtc.Clear();
            }
        }

        private void list_tc_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_tc.SelectedItem != null)
            {
                string selectedProduct = list_tc.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_mabtc.Text = selectedProduct;
                    list_tc.Visible = false;
                    string sql = "select HoTen from GiangVien where MaGV = '" + selectedProduct + "' ";
                    DataTable tb = my.DocDL(sql);
                    if (tb.Rows.Count > 0)
                    {
                        string hoten = tb.Rows[0][0].ToString();
                        txt_tenbtc.Text = hoten;
                    }

                }

            }
        }
        private void ShowSuggestionsHT(List<string> suggestions)
        {
            list_ht.Items.Clear();
            list_ht.Items.AddRange(suggestions.ToArray());

            list_ht.Visible = suggestions.Any();
        }
        private void txt_mabkt_TextChanged(object sender, EventArgs e)
        {
            string searchTerm = txt_mabkt.Text.ToLower();
            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                List<string> filteredProducts = productListGV
               .Where(product => product.ToLower().Contains(searchTerm))
               .ToList();

                if (filteredProducts != null)
                {
                    ShowSuggestionsHT(filteredProducts);
                }
                else
                {
                    list_ht.Visible = false;

                }


            }
            else
            {
                list_ht.Visible = false;
                txt_tenbkt.Clear();
            }
        }

        private void list_ht_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_ht.SelectedItem != null)
            {
                string selectedProduct = list_ht.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_mabkt.Text = selectedProduct;
                    list_ht.Visible = false;
                    string sql = "select HoTen from GiangVien where MaGV = '" + selectedProduct + "' ";
                    DataTable tb = my.DocDL(sql);
                    if (tb.Rows.Count > 0)
                    {
                        string hoten = tb.Rows[0][0].ToString();
                        txt_tenbkt.Text = hoten;
                    }

                }

            }
        }
    }
}