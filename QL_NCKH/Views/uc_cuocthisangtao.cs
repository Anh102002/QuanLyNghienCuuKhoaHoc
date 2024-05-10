using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Excel=Microsoft.Office.Interop.Excel;
using System.Threading;

namespace QL_NCKH
{
    public partial class uc_cuocthisangtao : DevExpress.XtraEditors.XtraUserControl
    {
        MyClass my = new MyClass();
        private List<string> productListGV;
        private List<string> productListSV;
        private string mact;
        public uc_cuocthisangtao()
        {
            InitializeComponent();
        }
        public string Mact
        {
            get { return this.mact; }
            set {  this.mact = value; }
        }
        public void loadDL()
        {
            try
            {
               
                string query = " select * from CuocThiSTKN where CapCT = N'Cấp Quốc Gia' or  CapCT = N'Cấp Bộ' ";
                DataTable dt = my.DocDL(query);
                dgv_ct.DataSource = dt;
                dgv_ct.Columns[0].HeaderText = "Mã cuộc thi";
                dgv_ct.Columns[1].HeaderText = "Tên cuộc thi";
                dgv_ct.Columns[1].Width = 300;
                dgv_ct.Columns[2].HeaderText = "Lĩnh vực";
                dgv_ct.Columns[3].HeaderText = "Năm tổ chức";
                dgv_ct.Columns[4].HeaderText = "Kinh phí";
                dgv_ct.Columns[5].HeaderText = "Loại cuộc thi";
                dgv_ct.Columns[6].HeaderText = "Cấp cuộc thi";
                txt_chucvugv.Text = "Người hướng dẫn";
               

               
            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu cuộc thi ","Lỗi");
            }
        }

        public void loadDLDT(string ma)
        {
            try
            {

                string query = " select MaDoi,TenDoi,TenYTuong,DonVi from DoiThamGiaCuocThi where MaCuocThi = '" + ma + "' ";
                DataTable dt = my.DocDL(query);
                dgv_doithi.DataSource = dt;
                dgv_doithi.Columns[0].HeaderText = "Mã đội thi";
                dgv_doithi.Columns[1].HeaderText = "Tên đội thi";
                dgv_doithi.Columns[2].HeaderText = "Tên ý tưởng";
                dgv_doithi.Columns[3].HeaderText = "Đơn vị";
                dgv_doithi.Columns[1].Width = 200;

            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu đội thi ", "Lỗi");
            }
        }

        public void loadDLTVDT(string ma)
        {
            try
            {

                string query = @" select ThanhVienCuocThi.MaSV,SinhVien.HoTen,ThanhVienCuocThi.ChucVu 
                                from ThanhVienCuocThi,SinhVien 
                                where ThanhVienCuocThi.MaSV = SinhVien.MaSV and ThanhVienCuocThi.MaDoi =  '" + ma + "' ";
                DataTable dt = my.DocDL(query);
                dgv_sv.DataSource = dt;
                dgv_sv.Columns[0].HeaderText = "Mã sinh viên";
                dgv_sv.Columns[1].HeaderText = "Tên sinh viên";
                dgv_sv.Columns[2].HeaderText = "Chức vụ";

            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu thành viên đội thi ", "Lỗi");
            }
        }
        public void loadDLGVHD(string ma)
        {
            try
            {

                string query = @" select GVHDCuocThi.MaGV,GiangVien.HoTen,GVHDCuocThi.ChucVu 
                                from GVHDCuocThi,GiangVien 
                                where GVHDCuocThi.MaGV = GiangVien.MaGV and GVHDCuocThi.MaDoi =  '" + ma + "' ";
                DataTable dt = my.DocDL(query);
                dgv_gv.DataSource = dt;
                dgv_gv.Columns[0].HeaderText = "Mã giảng viên";
                dgv_gv.Columns[1].HeaderText = "Tên giảng viên";
                dgv_gv.Columns[2].HeaderText = "Chức vụ";

            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu giáo viên hướng dẫn ", "Lỗi");
            }
        }
        private void uc_cuocthisangtao_Load(object sender, EventArgs e)
        {
            try
            {
                loadDL();
                LoadProductList();
                LoadProductListGV();
                txt_chucvugv.Text = "Người hướng dẫn";
            }
            catch
            {
                MessageBox.Show("Lỗi load dữ liệu ", "Lỗi");
            }

        }

        private void dgv_ct_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_mact.Text = dgv_ct.CurrentRow.Cells[0].Value.ToString();
                txt_tenct.Text = dgv_ct.CurrentRow.Cells[1].Value.ToString();
                txt_linhvuc.Text = dgv_ct.CurrentRow.Cells[2].Value.ToString();
                txt_nam.Text = dgv_ct.CurrentRow.Cells[3].Value.ToString();
                txt_kinhphi.Text = dgv_ct.CurrentRow.Cells[4].Value.ToString();
                cbo_loaict.Text = dgv_ct.CurrentRow.Cells[5].Value.ToString();
                cbo_capct.Text = dgv_ct.CurrentRow.Cells[6].Value.ToString();

                Mact = txt_mact.Text;
                string ma = Mact;
                loadDLDT(ma);
                dgv_sv.DataSource = null;
                dgv_gv.DataSource = null;
                
            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị thông tin cuộc thi","Lỗi");
            }
        }

        private void dgv_doithi_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_madoithi.Text = dgv_doithi.CurrentRow.Cells[0].Value.ToString();
                txt_tendoithi.Text = dgv_doithi.CurrentRow.Cells[1].Value.ToString();
                txt_tenytuong.Text = dgv_doithi.CurrentRow.Cells[2].Value.ToString();
                txt_donvi.Text = dgv_doithi.CurrentRow.Cells[3].Value.ToString();
                string ma = txt_madoithi.Text;
                loadDLTVDT(ma);
                loadDLGVHD(ma);
                
               
            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị thông tin đội thi", "Lỗi");
            }
        }

        private void dgv_sv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_masv.Text = dgv_sv.CurrentRow.Cells[0].Value.ToString();
                txt_tensv.Text = dgv_sv.CurrentRow.Cells[1].Value.ToString();
                cbo_chucvusv.Text = dgv_sv.CurrentRow.Cells[2].Value.ToString();
            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị thông tin đội thi", "Lỗi");
            }
        }

        private void dgv_gv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_magv.Text = dgv_gv.CurrentRow.Cells[0].Value.ToString();
                txt_tengv.Text = dgv_gv.CurrentRow.Cells[1].Value.ToString();
                txt_chucvugv.Text = dgv_gv.CurrentRow.Cells[2].Value.ToString();
            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị giáo viên hướng dẫn", "Lỗi");
            }
        }
        public bool Ktra()
        {
            if (string.IsNullOrWhiteSpace(txt_mact.Text) || string.IsNullOrWhiteSpace(txt_linhvuc.Text)
                || string.IsNullOrWhiteSpace(txt_tenct .Text) || string.IsNullOrWhiteSpace(txt_kinhphi.Text) 
                || string.IsNullOrWhiteSpace(txt_nam.Text) || string.IsNullOrWhiteSpace(cbo_loaict.Text) || string.IsNullOrWhiteSpace(cbo_capct.Text))
                return false;

            return true;
        }
        public bool KtraMaDT(string ma)
        {
            try
            {
                string sql = "select * from DoiThamGiaCuocThi where MaDoi = '" + ma + "'";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }


            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã đội thi !", "Thông báo");
            }
            return true;
        }
        public bool KtraMaCT(string ma)
        {
            try
            {
                string sql = "select * from CuocThiSTKN where MaCuocThi = '" + ma + "'";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }


            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã cuộc thi !", "Thông báo");
            }
            return true;
        }
        public bool kiemtraNam()
        {
            int nam;
            string namm = txt_nam.Text;
            if (!int.TryParse(namm, out nam) || namm.Count() > 4)
            {
                return false;
            }

            return true;
        }
        public bool kiemtraKinhphi()
        {
            decimal nam;
            string namm = txt_kinhphi.Text;
            if (!decimal.TryParse(namm, out nam))
            {
                return false;
            }

            return true;
        }

        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Ktra())
            {
                if (kiemtraNam())
                {
                    if (KtraMaCT(txt_mact.Text))
                    {
                        if (kiemtraKinhphi())
                        {
                            try
                            {

                                string sql = "insert into CuocThiSTKN values (@Ma,@Ten,@Linhvuc,@Nam,@Kinhphi,@Loai,@Cap) ";
                                SqlCommand command = my.SqlCommand(sql);
                                command.Parameters.AddWithValue("@Ma", txt_mact.Text);
                                command.Parameters.AddWithValue("@Ten", txt_tenct.Text);
                                command.Parameters.AddWithValue("@Linhvuc", txt_linhvuc.Text);
                                command.Parameters.AddWithValue("@Nam", txt_nam.Text);
                                command.Parameters.AddWithValue("@Kinhphi", txt_kinhphi.Text);
                                command.Parameters.AddWithValue("@Loai", cbo_loaict.Text);
                                command.Parameters.AddWithValue("@Cap", cbo_capct.Text);

                                int up = command.ExecuteNonQuery();
                                if (up > 0)
                                {
                                    MessageBox.Show("Thêm thông tin thành công", "Thông báo");
                                    txt_mact.Clear();
                                    txt_tenct.Clear();
                                    txt_linhvuc.Clear();
                                    txt_nam.Clear();
                                    txt_kinhphi.Clear();
                                    cbo_loaict.SelectedIndex = -1;

                                    loadDL();

                                }



                            }
                            catch
                            {
                                MessageBox.Show("Lỗi ! không thêm thành công ", "Thông báo");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Vui lòng nhập lại kinh phí !", "Thông báo");
                        }

                    }
                    else
                    {
                        MessageBox.Show("Đã có mã cuộc thi này !", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng nhập lại năm tổ chức !", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin cuộc thi !", "Thông báo");
            }
        }

        private void dgv_ct_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btn_sua_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Ktra())
            {
                if (kiemtraNam())
                {
                    if (!KtraMaCT(txt_mact.Text))
                    {
                        if (kiemtraKinhphi())
                        {
                            try
                            {

                                string sql = "update CuocThiSTKN set TenCuocThi=@Ten,LinhVuc=@Linhvuc,NamTC=@Nam,KinhPhi=@Kinhphi,LoaiCuocThi=@Loai,CapCT=@Cap where MaCuocThi = @Ma ";
                                SqlCommand command = my.SqlCommand(sql);
                                command.Parameters.AddWithValue("@Ma", txt_mact.Text);
                                command.Parameters.AddWithValue("@Ten", txt_tenct.Text);
                                command.Parameters.AddWithValue("@Linhvuc", txt_linhvuc.Text);
                                command.Parameters.AddWithValue("@Nam", txt_nam.Text);
                                command.Parameters.AddWithValue("@Kinhphi", txt_kinhphi.Text);
                                command.Parameters.AddWithValue("@Loai", cbo_loaict.Text);
                                command.Parameters.AddWithValue("@Cap",cbo_capct.Text);

                                int up = command.ExecuteNonQuery();
                                if (up > 0)
                                {
                                    MessageBox.Show("Sửa thông tin thành công", "Thông báo");
                                    txt_mact.Clear();
                                    txt_tenct.Clear();
                                    txt_linhvuc.Clear();
                                    txt_nam.Clear();
                                    txt_kinhphi.Clear();
                                    cbo_loaict.SelectedIndex = -1;

                                    loadDL();

                                }



                            }
                            catch
                            {
                                MessageBox.Show("Lỗi ! không sửa thành công ", "Thông báo");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Vui lòng nhập lại kinh phí !", "Thông báo");
                        }

                    }
                    else
                    {
                        MessageBox.Show("không có mã cuộc thi này !", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng nhập lại năm tổ chức !", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin cuộc thi !", "Thông báo");
            }
        }
        public List<string> LayMaDT()
        {
            try
            {
                List<string> ma = new List<string>();
               
                string sql = "select * from DoiThamGiaCuocThi where MaCuocThi = '" +txt_mact.Text+ "'";
                DataTable tb = my.DocDL(sql);
                for (int i = 0; i < tb.Rows.Count; i++)
                {
                    string maDT = tb.Rows[i]["MaDoi"].ToString();
                    ma.Add(maDT);
                }

                return ma;
            }
            catch
            {
                MessageBox.Show("Lỗi tạo danh sách mã đội thi !", "Thông báo");
            }
            return null;
        }
        private void btn_xoa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Ktra())
            {
                if (kiemtraNam())
                {
                    if (!KtraMaCT(txt_mact.Text))
                    {
                        if (kiemtraKinhphi())
                        {
                            try
                            {
                                DialogResult result = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                                if (result == DialogResult.OK)
                                {
                                    int count = 0;
                                    List<string> maDT = LayMaDT();
                                    if (maDT.Count > 0)
                                    {

                                        for (int i = 0; i < maDT.Count; i++)
                                        {
                                            string ma = maDT[i].ToString();
                                            string queryGV = "delete from GVHDCuocThi where MaDoi = @Madoi";
                                            SqlCommand commandGV = my.SqlCommand(queryGV);
                                            commandGV.Parameters.AddWithValue("@Madoi", ma);
                                            int xoaGV = commandGV.ExecuteNonQuery();
                                            if (xoaGV >= 0)
                                            {
                                                string querySV = "delete from ThanhVienCuocThi where MaDoi = @Madoi";
                                                SqlCommand commandSV = my.SqlCommand(querySV);
                                                commandSV.Parameters.AddWithValue("@Madoi", ma);
                                                int xoaSV = commandSV.ExecuteNonQuery();
                                                if (xoaSV >= 0)
                                                {
                                                    
                                                        string btt = "delete from BaiThuyetTrinhCT where MaDoi = @MaDoi ";

                                                        SqlCommand commandBTT = my.SqlCommand(btt);
                                                        commandBTT.Parameters.AddWithValue("@MaDoi", ma);

                                                        int upBTT = commandBTT.ExecuteNonQuery();
                                                        if (upBTT >= 0)
                                                        {

                                                            count++;
                                                        }
                                                    
                                                }
                                            }
                                        }
                                    }
                                    if (count > 0 || count == 0)
                                    {
                                        string queryDT = "delete from DoiThamGiaCuocThi where MaCuocThi = @Ma ";
                                        SqlCommand commandDT = my.SqlCommand(queryDT);
                                        commandDT.Parameters.AddWithValue("@Ma", txt_mact.Text);
                                        int xoaDT = commandDT.ExecuteNonQuery();
                                        if (xoaDT >= 0)
                                        {
                                            string bgk = "delete from BGKCuocThi where MaCuocThi=@Ma1 ";
                                            SqlCommand commandbgk = my.SqlCommand(bgk);
                                            commandbgk.Parameters.AddWithValue("@Ma1", txt_mact.Text);

                                            int upbgk = commandbgk.ExecuteNonQuery();
                                            if (upbgk >= 0)
                                            {
                                                string bgknt = "delete from BGKCuocThiKM where MaCuocThi=@Ma2 ";
                                                SqlCommand commandbgknt = my.SqlCommand(bgknt);
                                                commandbgknt.Parameters.AddWithValue("@Ma2", txt_mact.Text);

                                                int upbgknt = commandbgknt.ExecuteNonQuery();
                                                if (upbgknt >= 0)
                                                {
                                                    string sqlKQCT = "delete from KetQuaCuocThi where MaCuocThi=@Ma3 ";
                                                    SqlCommand commandKQCT = my.SqlCommand(sqlKQCT);
                                                    commandKQCT.Parameters.AddWithValue("@Ma3", txt_mact.Text);



                                                    int upKQCT = commandKQCT.ExecuteNonQuery();
                                                    if (upKQCT >= 0)
                                                    {
                                                        string sqlLD = "delete from BanLanhDaoCT where MaCuocThi=@Ma";
                                                        SqlCommand commandLD = my.SqlCommand(sqlLD);
                                                        commandLD.Parameters.AddWithValue("@Ma", txt_mact.Text);
                                                        int upLD = commandLD.ExecuteNonQuery();
                                                        if (upLD >= 0)
                                                        {
                                                            string sqlLDNT = "delete from BanLanhDaoCTNT where MaCuocThi=@Ma";
                                                            SqlCommand commandLDNT = my.SqlCommand(sqlLDNT);
                                                            commandLDNT.Parameters.AddWithValue("@Ma", txt_mact.Text);
                                                            int upLDNT = commandLDNT.ExecuteNonQuery();
                                                            if (upLDNT >= 0)
                                                            {
                                                                string sqlTC = "delete from BanToChucCT where MaCuocThi=@Ma";
                                                                SqlCommand commandTC = my.SqlCommand(sqlTC);
                                                                commandTC.Parameters.AddWithValue("@Ma", txt_mact.Text);
                                                                int upTC = commandTC.ExecuteNonQuery();
                                                                if (upTC >= 0)
                                                                {
                                                                    string sqlHT = "delete from BanHoTroCT where MaCuocThi=@Ma";
                                                                    SqlCommand commandHT = my.SqlCommand(sqlHT);
                                                                    commandHT.Parameters.AddWithValue("@Ma", txt_mact.Text);
                                                                    int upHT = commandHT.ExecuteNonQuery();
                                                                    if (upHT >= 0)
                                                                    {
                                                                        string sql = "delete from CuocThiSTKN where MaCuocThi = @Mact ";
                                                                        SqlCommand command = my.SqlCommand(sql);
                                                                        command.Parameters.AddWithValue("@Mact", txt_mact.Text);


                                                                        int up = command.ExecuteNonQuery();
                                                                        if (up > 0)
                                                                        {
                                                                            MessageBox.Show("Xóa thông tin thành công", "Thông báo");
                                                                            txt_mact.Clear();
                                                                            txt_tenct.Clear();
                                                                            txt_linhvuc.Clear();
                                                                            txt_nam.Clear();
                                                                            txt_kinhphi.Clear();
                                                                            cbo_loaict.SelectedIndex = -1;

                                                                            loadDL();
                                                                            txt_madoithi.Clear();
                                                                            txt_tendoithi.Clear();
                                                                            dgv_doithi.DataSource = null;
                                                                            txt_masv.Clear();
                                                                            txt_tensv.Clear();
                                                                            cbo_chucvusv.SelectedIndex = -1;
                                                                            dgv_sv.DataSource = null;
                                                                            txt_tengv.Clear();
                                                                            txt_magv.Clear();
                                                                            txt_chucvugv.Clear();
                                                                            dgv_gv.DataSource = null;

                                                                            
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
                            catch (Exception ex)
                            {
                                MessageBox.Show("Lỗi ! không xóa thành công {" + ex.Message + "}", "Thông báo");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Vui lòng nhập lại kinh phí !", "Thông báo");
                        }

                    }
                    else
                    {
                        MessageBox.Show("không có mã cuộc thi này !", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng nhập lại năm tổ chức !", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Vui chọn cuộc thi muốn xóa !", "Thông báo");
            }
        }
        public bool KtraMaDoiThi(string madt)
        {
            try
            {
                string sql = "select * from DoiThamGiaCuocThi where MaDoi = '" + madt + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }
            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã đội thi !", "Thông báo");
            }
            return true;
        }

        public bool KtraMaTVDT(string madt, string masv)
        {
            try
            {
                string sql = "select * from ThanhVienCuocThi where MaDoi = '" + madt + "' and MaSV = '" + masv + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }
            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã thành viên !", "Thông báo");
            }
            return true;
        }

        private void btn_joingv_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrWhiteSpace(txt_madoithi.Text) || String.IsNullOrWhiteSpace(txt_tendoithi.Text)
                        || String.IsNullOrWhiteSpace(txt_tenytuong.Text) || string.IsNullOrWhiteSpace(txt_donvi.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {

                    if (string.IsNullOrWhiteSpace(Mact) || dgv_doithi.DataSource == null)
                    {
                        MessageBox.Show("Vui lòng chọn cuộc thi !", "Thông báo");
                    }
                    else
                    {
                        string mact = Mact;
                        if (KtraMaDT(txt_madoithi.Text))
                        {


                            string sql = "insert into DoiThamGiaCuocThi values (@Madt,@Mact,@Tendt,@TenYtuong,@DonVi,@Giaithuong)";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Madt", txt_madoithi.Text);
                            comd.Parameters.AddWithValue("@Mact", mact);
                            comd.Parameters.AddWithValue("@Tendt", txt_tendoithi.Text);
                            comd.Parameters.AddWithValue("@TenYtuong", txt_tenytuong.Text);
                            comd.Parameters.AddWithValue("@DonVi", txt_donvi.Text);
                            comd.Parameters.AddWithValue("@Giaithuong", "");

                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Thêm đội thi thành công !", "Thông báo");
                                txt_madoithi.Clear();
                                txt_tendoithi.Clear();
                                txt_donvi.Clear();
                                txt_tenytuong.Clear();
                                loadDLDT(mact);
                            }
                            else
                            {
                                MessageBox.Show("Thêm đội thi không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("Đội thi đã có trong cuộc thi !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi thêm đội thi!", "Thông báo");
                }
            }
        }

        private void btn_cancelgv_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrWhiteSpace(txt_madoithi.Text) || String.IsNullOrWhiteSpace(txt_tendoithi.Text)
                         || String.IsNullOrWhiteSpace(txt_tenytuong.Text) || string.IsNullOrWhiteSpace(txt_donvi.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(Mact) || dgv_doithi.DataSource == null)
                    {
                        MessageBox.Show("Vui lòng chọn cuộc thi !", "Thông báo");
                    }
                    else
                    {
                        string mact = Mact;
                        if (!KtraMaDT(txt_madoithi.Text))
                        {

                            string sql = "delete from ThanhVienCuocThi where  MaDoi = @Madt ";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Madt", txt_madoithi.Text);



                            int up = comd.ExecuteNonQuery();
                            if (up >= 0)
                            {

                                string gv = "delete from GVHDCuocThi where MaDoi = @Madt ";
                                SqlCommand comdGV = my.SqlCommand(gv);
                                comdGV.Parameters.AddWithValue("@Madt", txt_madoithi.Text);



                                int upGV = comdGV.ExecuteNonQuery();
                                if (upGV >= 0)
                                {
                                    string svnt = "delete from ThanhVienCuocThiNT where  MaDoi = @Madt ";
                                    SqlCommand comdSVNT = my.SqlCommand(svnt);
                                    comdSVNT.Parameters.AddWithValue("@Madt", txt_madoithi.Text);
                                    int upSVNT = comdSVNT.ExecuteNonQuery();
                                    if (upSVNT >= 0)
                                    {
                                        string dt = "delete from DoiThamGiaCuocThi where  MaDoi = @Madt ";
                                        SqlCommand comdDT = my.SqlCommand(dt);
                                        comdDT.Parameters.AddWithValue("@Madt", txt_madoithi.Text);

                                        int upDT = comdDT.ExecuteNonQuery();
                                        if (upDT > 0)
                                        {

                                            MessageBox.Show("Xóa đội thi thành công !", "Thông báo");
                                            txt_madoithi.Clear();
                                            txt_tendoithi.Clear();
                                            txt_donvi.Clear();
                                            txt_tenytuong.Clear();
                                            txt_tensv.Clear();
                                            txt_masv.Clear();
                                            cbo_chucvusv.SelectedIndex = -1;
                                            dgv_sv.DataSource = null;
                                           // txt_tensvnt.Clear();
                                           // txt_masvnt.Clear();
                                           // cbo_chucvunt.SelectedIndex = -1;
                                           // dgv_svnt.DataSource = null;
                                            txt_tengv.Clear();
                                            txt_magv.Clear();
                                            dgv_gv.DataSource = null;
                                            loadDLDT(mact);
                                        }
                                        else
                                        {
                                            MessageBox.Show("Xóa đội thi không thành công !", "Thông báo");
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Xóa đội thi không thành công !", "Thông báo");
                                    }


                                }
                                else
                                {
                                    MessageBox.Show("Xóa đội thi không thành công !", "Thông báo");
                                }
                            }
                            else
                            {
                                MessageBox.Show("Xóa đội thi không thành công !", "Thông báo");
                            }








                        }
                        else
                        {
                            MessageBox.Show("Đội thi không có trong cuộc thi !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi xóa đội thi!", "Thông báo");
                }
            }
        }

        private void btn_suadoithi_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_madoithi.Text) || String.IsNullOrWhiteSpace(txt_tendoithi.Text)
                       || String.IsNullOrWhiteSpace(txt_tenytuong.Text) || string.IsNullOrWhiteSpace(txt_donvi.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(Mact) || dgv_doithi.DataSource == null)
                    {
                        MessageBox.Show("Vui lòng chọn cuộc thi !", "Thông báo");
                    }
                    else
                    {
                        string mact = Mact;
                        if (!KtraMaDT(txt_madoithi.Text))
                        {


                            string sql = "update DoiThamGiaCuocThi set TenDoi= @Tendoi,TenYTuong = @TenYtuong,DonVi=@DonVi where  MaDoi = @Madt and MaCuocThi =@Mact ";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Madt", txt_madoithi.Text);
                            comd.Parameters.AddWithValue("@Mact", mact);
                            comd.Parameters.AddWithValue("@Tendoi", txt_tendoithi.Text);
                            comd.Parameters.AddWithValue("@TenYtuong", txt_tenytuong.Text);
                            comd.Parameters.AddWithValue("@DonVi", txt_donvi.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa đội thi thành công !", "Thông báo");
                                txt_madoithi.Clear();
                                txt_tendoithi.Clear();
                                txt_donvi.Clear();
                                txt_tenytuong.Clear();
                                loadDLDT(mact);
                            }
                            else
                            {
                                MessageBox.Show("Sửa đội thi không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("Đội thi không có trong cuộc thi !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi sửa đội thi!", "Thông báo");
                }
            }
        }

        private void btn_joinhd_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_masv.Text) || string.IsNullOrWhiteSpace(txt_tensv.Text) || string.IsNullOrWhiteSpace(cbo_chucvusv.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_madoithi.Text) || String.IsNullOrWhiteSpace(txt_tendoithi.Text))
                    {
                        MessageBox.Show("Vui lòng chọn đội thi !", "Thông báo");
                    }
                    else
                    {
                        string mact = Mact;
                        if (KtraMaTVDT(txt_madoithi.Text, txt_masv.Text))
                        {
                            string sql = "insert into ThanhVienCuocThi values (@Masv,@Madoi,@ChucVu)";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Masv", txt_masv.Text);
                            comd.Parameters.AddWithValue("@Madoi", txt_madoithi.Text);
                            comd.Parameters.AddWithValue("@ChucVu", cbo_chucvusv.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Thêm thành viên thành công !", "Thông báo");

                                txt_tensv.Clear();
                                txt_masv.Clear();
                                cbo_chucvusv.SelectedIndex = -1;
                                loadDLTVDT(txt_madoithi.Text);

                            }
                            else
                            {
                                MessageBox.Show("Thêm thành viên không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("Đội thi đã có trong cuộc thi !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi thêm đội thi!", "Thông báo");
                }
            }
        }

        private void btn_cancelhd_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_masv.Text) || string.IsNullOrWhiteSpace(txt_tensv.Text) || string.IsNullOrWhiteSpace(cbo_chucvusv.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_madoithi.Text) || String.IsNullOrWhiteSpace(txt_tendoithi.Text))
                    {
                        MessageBox.Show("Vui lòng chọn đội thi !", "Thông báo");
                    }
                    else
                    {
                        string mact = Mact;
                        if (!KtraMaTVDT(txt_madoithi.Text, txt_masv.Text))
                        {
                            string sql = "delete from ThanhVienCuocThi where MaSV = @Masv and MaDoi = @Madoi ";
                            SqlCommand comd = my.SqlCommand(sql);

                            comd.Parameters.AddWithValue("@Masv", txt_masv.Text);
                            comd.Parameters.AddWithValue("@Madoi", txt_madoithi.Text);



                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa thành viên thành công !", "Thông báo");

                                txt_tensv.Clear();
                                txt_masv.Clear();
                                cbo_chucvusv.SelectedIndex = -1;
                                loadDLTVDT(txt_madoithi.Text);

                            }
                            else
                            {
                                MessageBox.Show("Xóa thành viên không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("Thành viên không có trong đội thi !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi xóa thành viên!", "Thông báo");
                }
            }
        }

        private void btn_suahd_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_masv.Text) || string.IsNullOrWhiteSpace(txt_tensv.Text) || string.IsNullOrWhiteSpace(cbo_chucvusv.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_madoithi.Text) || String.IsNullOrWhiteSpace(txt_tendoithi.Text))
                    {
                        MessageBox.Show("Vui lòng chọn đội thi !", "Thông báo");
                    }
                    else
                    {
                        string mact = Mact;
                        if (!KtraMaTVDT(txt_madoithi.Text, txt_masv.Text))
                        {
                            string sql = "update ThanhVienCuocThi set ChucVu=@Chucvu where MaSV = @Masv and MaDoi = @Madoi ";
                            SqlCommand comd = my.SqlCommand(sql);

                            comd.Parameters.AddWithValue("@Masv", txt_masv.Text);
                            comd.Parameters.AddWithValue("@Madoi", txt_madoithi.Text);
                            comd.Parameters.AddWithValue("@Chucvu", cbo_chucvusv.Text);



                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa thành viên thành công !", "Thông báo");

                                txt_tensv.Clear();
                                txt_masv.Clear();
                                cbo_chucvusv.SelectedIndex = -1;
                                loadDLTVDT(txt_madoithi.Text);

                            }
                            else
                            {
                                MessageBox.Show("Sửa thành viên không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("Thành viên không có trong đội thi !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi sửa thành viên!", "Thông báo");
                }
            }
        }
        private void LoadProductList()
        {
            try
            {

                productListSV = new List<string>();
                string query = "SELECT MaSV FROM SinhVien";
                DataTable tb = my.DocDL(query);
                if (tb.Rows.Count > 0)
                {
                    for (int i = 0; i < tb.Rows.Count; i++)
                    {
                        string ma = tb.Rows[i][0].ToString();
                        productListSV.Add(ma);
                    }
                }
            }
            catch
            {
                MessageBox.Show($"Lỗi thực hiện tạo danh sách sinh viên", "Lỗi");
            }

        }
        private void ShowSuggestions(List<string> suggestions)
        {
            list_sv.Items.Clear();
            list_sv.Items.AddRange(suggestions.ToArray());

            list_sv.Visible = suggestions.Any();
        }
        private void cbo_masv_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void list_sv_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_sv.SelectedItem != null)
            {
                string selectedProduct = list_sv.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_masv.Text = selectedProduct;
                    list_sv.Visible = false;
                    string sql = "select HoTen from SinhVien where MaSV = '" + selectedProduct + "' ";
                    DataTable tb = my.DocDL(sql);
                    if (tb.Rows.Count > 0)
                    {
                        string hoten = tb.Rows[0][0].ToString();
                        txt_tensv.Text = hoten;
                    }
                    
                }

            }
        }

        private void cbo_masv_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
                string searchTerm = txt_masv.Text.ToLower();
                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
                    List<string> filteredProducts = productListSV
                   .Where(product => product.ToLower().Contains(searchTerm))
                   .ToList();

                    if(filteredProducts != null)
                    {
                        ShowSuggestions(filteredProducts);
                    }
                    else
                    {
                        list_sv.Visible = false;
                        
                    }
                    

                }
                else
                {
                    list_sv.Visible = false;
                    txt_tensv.Clear();
                }


            
            
            
        }

        private void btn_refresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            cbo_tk.SelectedIndex = -1;
            txt_timkiem.Clear();
            txt_mact.Clear();
            txt_tenct.Clear();
            txt_linhvuc.Clear();
            txt_nam.Clear();
            txt_kinhphi.Clear();
            cbo_loaict.SelectedIndex = -1;
            cbo_capct.SelectedIndex = -1;
            loadDL();
            txt_madoithi.Clear();
            txt_tendoithi.Clear();
            dgv_doithi.DataSource = null;
            txt_masv.Clear();
            txt_tensv.Clear();
            cbo_chucvusv.SelectedIndex = -1;
            dgv_sv.DataSource = null;
            txt_tengv.Clear();
            txt_magv.Clear();
            txt_chucvugv.Clear();
            dgv_gv.DataSource = null;
            txt_chucvugv.Text = "Người hướng dẫn";
            txt_tenytuong.Clear();
            txt_donvi.Clear();
        }

        private void btn_timkiem_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
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
                   
                    if (cbo_tk.SelectedIndex == 0)
                    {
                        try
                        {
                            string query = " select * from CuocThiSTKN where MaCuocThi like '%" + txt_timkiem.Text + "%' and CapCT = N'Cấp Quốc Gia' or  CapCT = N'Cấp Bộ'  ";
                            DataTable dt = my.DocDL(query);
                            dgv_ct.DataSource = dt;
                            dgv_ct.Columns[0].HeaderText = "Mã cuộc thi";
                            dgv_ct.Columns[1].HeaderText = "Tên cuộc thi";
                            dgv_ct.Columns[1].Width = 300;
                            dgv_ct.Columns[2].HeaderText = "Lĩnh vực";
                            dgv_ct.Columns[3].HeaderText = "Năm tổ chức";
                            dgv_ct.Columns[4].HeaderText = "Kinh phí";
                            dgv_ct.Columns[5].HeaderText = "Loại cuộc thi";
                            dgv_ct.Columns[6].HeaderText = "Cấp cuộc thi";


                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm theo mã cuộc thi  !", "Thông báo");
                        }
                    }
                    else if (cbo_tk.SelectedIndex == 1)
                    {
                        try
                        {

                            string query = " select * from CuocThiSTKN where CapCT = N'Cấp Quốc Gia' or  CapCT = N'Cấp Bộ' and TenCuocThi like N'%" + txt_timkiem.Text + "%' ";
                            DataTable dt = my.DocDL(query);
                            dgv_ct.DataSource = dt;
                            dgv_ct.Columns[0].HeaderText = "Mã cuộc thi";
                            dgv_ct.Columns[1].HeaderText = "Tên cuộc thi";
                            dgv_ct.Columns[1].Width = 300;
                            dgv_ct.Columns[2].HeaderText = "Lĩnh vực";
                            dgv_ct.Columns[3].HeaderText = "Năm tổ chức";
                            dgv_ct.Columns[4].HeaderText = "Kinh phí";
                            dgv_ct.Columns[5].HeaderText = "Loại cuộc thi";
                            dgv_ct.Columns[6].HeaderText = "Cấp cuộc thi";

                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm theo tên cuộc thi !", "Thông báo");
                        }
                    }
                    

                }
            }
        }

        private void cbo_magv_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cbo_magv_TextChanged(object sender, EventArgs e)
        {
            
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
            list_gv.Items.Clear();
            list_gv.Items.AddRange(suggestions.ToArray());

            list_gv.Visible = suggestions.Any();
        }
        private void list_gv_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_gv.SelectedItem != null)
            {
                string selectedProduct = list_gv.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_magv.Text = selectedProduct;
                    list_gv.Visible = false;
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

        private void textBox1_TextChanged_1(object sender, EventArgs e)
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
                        list_gv.Visible = false;

                    }


                }
                else
                {
                    list_gv.Visible = false;
                    txt_tengv.Clear();
                }


            
            
        }
        public bool KtraMaGVHD(string madt, string magv)
        {
            try
            {
                string sql = "select * from GVHDCuocThi where MaDoi = '" + madt + "' and MaGV = '" + magv + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }
            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã giảng viên !", "Thông báo");
            }
            return true;
        }
        private void btn_joinhdnt_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrWhiteSpace(txt_magv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text) || string.IsNullOrWhiteSpace(txt_chucvugv.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_madoithi.Text) || String.IsNullOrWhiteSpace(txt_tendoithi.Text))
                    {
                        MessageBox.Show("Vui lòng chọn đội thi !", "Thông báo");
                    }
                    else
                    {
                        string mact = Mact;
                        if (KtraMaGVHD(txt_madoithi.Text, txt_magv.Text))
                        {
                            string sql = "insert into GVHDCuocThi values (@Magv,@Madoi,@ChucVu)";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Magv", txt_magv.Text);
                            comd.Parameters.AddWithValue("@Madoi", txt_madoithi.Text);
                            comd.Parameters.AddWithValue("@ChucVu", txt_chucvugv.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Thêm GVHD thành công !", "Thông báo");

                                txt_tengv.Clear();
                                txt_magv.Clear();

                                loadDLGVHD(txt_madoithi.Text);

                            }
                            else
                            {
                                MessageBox.Show("Thêm GVHD không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("GVHD đã có trong đội thi !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi thêm GVHD!", "Thông báo");
                }
            }
        }

        private void btn_cancelhdnt_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_magv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text) || string.IsNullOrWhiteSpace(txt_chucvugv.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_madoithi.Text) || String.IsNullOrWhiteSpace(txt_tendoithi.Text))
                    {
                        MessageBox.Show("Vui lòng chọn đội thi !", "Thông báo");
                    }
                    else
                    {
                        string mact = Mact;
                        if (!KtraMaGVHD(txt_madoithi.Text, txt_magv.Text))
                        {
                            string sql = "delete from GVHDCuocThi where MaGV = @Magv and MaDoi=@Madoi ";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Magv", txt_magv.Text);
                            comd.Parameters.AddWithValue("@Madoi", txt_madoithi.Text);



                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa GVHD thành công !", "Thông báo");

                                txt_tengv.Clear();
                                txt_magv.Clear();

                                loadDLGVHD(txt_madoithi.Text);


                            }
                            else
                            {
                                MessageBox.Show("Xóa GVHD không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("GVHD không có trong đội thi !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi xóa GVHD!", "Thông báo");
                }
            }
        }
        public void ExcelExportDSCT()
        {
            try
            {
                Excel.Application oExcel = new Excel.Application();
                Excel.Workbook oBook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

                Excel.Worksheet oSheet = (Excel.Worksheet)oBook.Worksheets[1];

                Excel.Range head = oSheet.get_Range("A1", "G1");

                head.MergeCells = true;

                head.Value2 = "DANH SÁCH CUỘC THI CẤP BỘ,QUỐC GIA";

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

                Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                cl6.Value = "Loại cuộc thi";

                Excel.Range cl7 = oSheet.get_Range("G3", "G3");
                cl7.Value = "Cấp cuộc thi";



                Excel.Range rowHead = oSheet.get_Range("A3", "G3");
                rowHead.Font.Bold = true;
                rowHead.Font.Size = 13;
                rowHead.Font.Name = "Times New Roman";
                rowHead.Borders.LineStyle = Excel.Constants.xlSolid;
                rowHead.Interior.ColorIndex = 6;
                rowHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int line = 4;
                for (int i = 0; i < dgv_ct.Rows.Count - 1; i++)
                {
                    Excel.Range line1 = oSheet.get_Range("A" + (line + i).ToString(), "A" + (line + i).ToString());
                    line1.Value = dgv_ct.Rows[i].Cells[0].Value.ToString();
                    line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line1.Borders.LineStyle = Excel.Constants.xlSolid;
                    line1.Font.Name = "Times New Roman";

                    Excel.Range line2 = oSheet.get_Range("B" + (line + i).ToString(), "B" + (line + i).ToString());
                    line2.Value = dgv_ct.Rows[i].Cells[1].Value.ToString();
                    line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line2.Borders.LineStyle = Excel.Constants.xlSolid;
                    line2.Font.Name = "Times New Roman";

                    Excel.Range line3 = oSheet.get_Range("C" + (line + i).ToString(), "C" + (line + i).ToString());
                    line3.Value = dgv_ct.Rows[i].Cells[2].Value.ToString();
                    line3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line3.Borders.LineStyle = Excel.Constants.xlSolid;
                    line3.Font.Name = "Times New Roman";

                    Excel.Range line4 = oSheet.get_Range("D" + (line + i).ToString(), "D" + (line + i).ToString());
                    line4.Value = dgv_ct.Rows[i].Cells[3].Value.ToString();
                    line4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line4.Borders.LineStyle = Excel.Constants.xlSolid;
                    line4.Font.Name = "Times New Roman";


                    Excel.Range line5 = oSheet.get_Range("E" + (line + i).ToString(), "E" + (line + i).ToString());
                    line5.Value = dgv_ct.Rows[i].Cells[4].Value.ToString();
                    line5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line5.Borders.LineStyle = Excel.Constants.xlSolid;
                    line5.Font.Name = "Times New Roman";

                    Excel.Range line6 = oSheet.get_Range("F" + (line + i).ToString(), "F" + (line + i).ToString());
                    line6.Value = dgv_ct.Rows[i].Cells[5].Value.ToString();
                    line6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line6.Borders.LineStyle = Excel.Constants.xlSolid;
                    line6.Font.Name = "Times New Roman";

                    Excel.Range line7 = oSheet.get_Range("G" + (line + i).ToString(), "G" + (line + i).ToString());
                    line7.Value = dgv_ct.Rows[i].Cells[6].Value.ToString();
                    line7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line7.Borders.LineStyle = Excel.Constants.xlSolid;
                    line7.Font.Name = "Times New Roman";

                }


                oSheet.Name = "CTSTKN";
                oExcel.Columns.AutoFit();

                oBook.Activate();

                SaveFileDialog saveFile = new SaveFileDialog();
                if (saveFile.ShowDialog() == DialogResult.OK)
                {

                    saveFile.Filter = "Các loại tập tin (*.xlsx;*.csv;*.docx)|*.xlsx;*.csv;*.docx|Tất cả các tập tin (*.*)|*.*";
                    oBook.SaveAs(saveFile.FileName.ToLower());
                    MessageBox.Show("Xuất danh sách thành công", "Thông báo");

                }

                oExcel.Quit();

            }
            catch
            {
                MessageBox.Show("Xuất danh sách không thành công");
            }
        }
        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ExcelExportDSCT();
        }

        public DataTable LayDuLieuBaoCao()
        {

            string query = " select * from CuocThiSTKN where CapCT = N'Cấp Quốc Gia' or  CapCT = N'Cấp Bộ' ";
            DataTable dataTable = my.DocDL(query);

            return dataTable;
        }
        public void excelCT()
        {
            try
            {
                
                DataTable dataTable = LayDuLieuBaoCao();


                Excel.Application oExcel = new Excel.Application();
                Excel.Workbook workbook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet oSheet = (Excel.Worksheet)workbook.Worksheets[1];




                Excel.Range head = oSheet.get_Range("A1", "J1");

                head.MergeCells = true;

                head.Value2 = "DANH SÁCH CHI TIẾT CÁC CUỘC THI CẤP BỘ,QUỐC GIA  ";

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
                cl6.Value = "Cấp cuộc thi";

                Excel.Range cl7 = oSheet.get_Range("H3", "H3");
                cl7.Value = "Đội thi";

                Excel.Range cl8 = oSheet.get_Range("I3", "I3");
                cl8.Value = "Thành viên đội thi - Chức vụ";

                Excel.Range cl9 = oSheet.get_Range("J3", "J3");
                cl9.Value = "Giảng viên hướng dẫn";



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
                    string query = "select MaDoi,TenDoi from DoiThamGiaCuocThi where MaCuocThi = '"+maCT+"' ";

                    DataTable dt = my.DocDL(query);

                    Excel.Range line1 = oSheet.get_Range("H" + (lines).ToString(), "H" + (lines).ToString());
                    Excel.Range line2 = oSheet.get_Range("I" + (lines).ToString(), "I" + (lines).ToString());
                    Excel.Range line3 = oSheet.get_Range("J" + (lines).ToString(), "J" + (lines).ToString());
                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        string maDT = dt.Rows[row][0].ToString();

                        string cel = dt.Rows[row]["MaDoi"].ToString() + "-" + dt.Rows[row]["TenDoi"].ToString() + "\n";
                        line1.Value += cel;
                        //
                        string hd = @"SELECT ThanhVienCuocThi.MaDoi,SinhVien.HoTen, ThanhVienCuocThi.ChucVu 
                                        FROM ThanhVienCuocThi,SinhVien 
                                        WHERE ThanhVienCuocThi.MaSV = SinhVien.MaSV AND ThanhVienCuocThi.MaDoi = '" + maDT + "' ";

                        DataTable tvdt = my.DocDL(hd);

                        

                        for (int r = 0; r < tvdt.Rows.Count; r++)
                        {

                            string celSV = tvdt.Rows[r]["MaDoi"].ToString() + "-" + tvdt.Rows[r]["HoTen"].ToString() + "-"+ tvdt.Rows[r]["ChucVu"].ToString()+"\n";
                            line2.Value += celSV;


                        }
                        line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line2.Borders.LineStyle = Excel.Constants.xlSolid;
                        line2.Font.Name = "Times New Roman";

                        //
                        string hdnt = @"SELECT GVHDCuocThi.MaDoi, GiangVien.HoTen 
                                        FROM GVHDCuocThi, GiangVien 
                                        WHERE GVHDCuocThi.MaGV = GiangVien.MaGV AND GVHDCuocThi.MaDoi = '" + maDT + "' ";
                        DataTable gvhd = my.DocDL(hdnt);

                        

                        for (int gv = 0; gv < gvhd.Rows.Count; gv++)
                        {

                            string celGV = gvhd.Rows[gv]["MaDoi"].ToString() + "-" + gvhd.Rows[gv]["HoTen"].ToString() + "\n";
                            line3.Value += celGV;


                        }
                        line3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line3.Borders.LineStyle = Excel.Constants.xlSolid;
                        line3.Font.Name = "Times New Roman";

                        //
                    }
                    line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line1.Borders.LineStyle = Excel.Constants.xlSolid;
                    line1.Font.Name = "Times New Roman";
                   //
                    lines++;




                }

                oSheet.Name = "CTCTSTKN";
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
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi xuất báo cáo: {ex.Message}");
            }
        }
        private void StartLongTask(frm_please formWaiting)
        {
            Thread.Sleep(3000);
            formWaiting.Invoke(new Action(() => formWaiting.Close()));
        }

        private void barButtonItem3_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            frm_please formWaiting = new frm_please();
            formWaiting.StartPosition = FormStartPosition.CenterScreen;

            Thread thread = new Thread(() => StartLongTask(formWaiting));
            thread.Start();

            formWaiting.ShowDialog();
            excelCT();
        }
        public DataTable LayDuLieuBaoCao1CT()
        {

            string query = " select * from CuocThiSTKN where MaCuocThi ='" + txt_mact.Text+"' ";
            DataTable dataTable = my.DocDL(query);

            return dataTable;
        }
        public void excelCT1CT()
        {
            try
            {
                if (Ktra())
                {
                    DataTable dataTable = LayDuLieuBaoCao1CT();
                    Excel.Application oExcel = new Excel.Application();
                    Excel.Workbook workbook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    Excel.Worksheet oSheet = (Excel.Worksheet)workbook.Worksheets[1];




                    Excel.Range head = oSheet.get_Range("A1", "J1");

                    head.MergeCells = true;

                    head.Value2 = "THÔNG TIN CUỘC THI CẤP BỘ,QUỐC GIA  ";

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
                    cl6.Value = "Cấp cuộc thi";

                    Excel.Range cl7 = oSheet.get_Range("H3", "H3");
                    cl7.Value = "Đội thi";

                    Excel.Range cl8 = oSheet.get_Range("I3", "I3");
                    cl8.Value = "Thành viên đội thi - Chức vụ";

                    Excel.Range cl9 = oSheet.get_Range("J3", "J3");
                    cl9.Value = "Giảng viên hướng dẫn";



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
                        string query = "select MaDoi,TenDoi from DoiThamGiaCuocThi where MaCuocThi = '" + maCT + "' ";

                        DataTable dt = my.DocDL(query);

                        Excel.Range line1 = oSheet.get_Range("H" + (lines).ToString(), "H" + (lines).ToString());
                        Excel.Range line2 = oSheet.get_Range("I" + (lines).ToString(), "I" + (lines).ToString());
                        Excel.Range line3 = oSheet.get_Range("J" + (lines).ToString(), "J" + (lines).ToString());
                        for (int row = 0; row < dt.Rows.Count; row++)
                        {
                            string maDT = dt.Rows[row][0].ToString();

                            string cel = dt.Rows[row]["MaDoi"].ToString() + "-" + dt.Rows[row]["TenDoi"].ToString() + "\n";
                            line1.Value += cel;
                            //
                            string hd = @"SELECT ThanhVienCuocThi.MaDoi,SinhVien.HoTen, ThanhVienCuocThi.ChucVu 
                                        FROM ThanhVienCuocThi,SinhVien 
                                        WHERE ThanhVienCuocThi.MaSV = SinhVien.MaSV AND ThanhVienCuocThi.MaDoi = '" + maDT + "' ";

                            DataTable tvdt = my.DocDL(hd);



                            for (int r = 0; r < tvdt.Rows.Count; r++)
                            {

                                string celSV = tvdt.Rows[r]["MaDoi"].ToString() + "-" + tvdt.Rows[r]["HoTen"].ToString() + "-" + tvdt.Rows[r]["ChucVu"].ToString() + "\n";
                                line2.Value += celSV;


                            }
                            line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            line2.Borders.LineStyle = Excel.Constants.xlSolid;
                            line2.Font.Name = "Times New Roman";

                            //
                            string hdnt = @"SELECT GVHDCuocThi.MaDoi, GiangVien.HoTen 
                                        FROM GVHDCuocThi, GiangVien 
                                        WHERE GVHDCuocThi.MaGV = GiangVien.MaGV AND GVHDCuocThi.MaDoi = '" + maDT + "' ";
                            DataTable gvhd = my.DocDL(hdnt);



                            for (int gv = 0; gv < gvhd.Rows.Count; gv++)
                            {

                                string celGV = gvhd.Rows[gv]["MaDoi"].ToString() + "-" + gvhd.Rows[gv]["HoTen"].ToString() + "\n";
                                line3.Value += celGV;


                            }
                            line3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            line3.Borders.LineStyle = Excel.Constants.xlSolid;
                            line3.Font.Name = "Times New Roman";

                            //
                        }
                        line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line1.Borders.LineStyle = Excel.Constants.xlSolid;
                        line1.Font.Name = "Times New Roman";
                        //
                        lines++;




                    }

                    oSheet.Name = "CTCTSTKN";
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
                else
                {
                    MessageBox.Show("Vui lòng chọn cuộc thi muốn export dữ liệu","Thông báo");
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi xuất báo cáo: {ex.Message}");
            }
        
                


                
        }
        private void barButtonItem4_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            excelCT1CT();
        }

        private void groupBox20_Enter(object sender, EventArgs e)
        {

        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {

        }

        private void button_tk_Click(object sender, EventArgs e)
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

                    if (cbo_tk.SelectedIndex == 0)
                    {
                        try
                        {
                            string query = " select * from CuocThiSTKN where MaCuocThi like '%" + txt_timkiem.Text + "%' and CapCT = N'Cấp Quốc Gia' or  CapCT = N'Cấp Bộ'  ";
                            DataTable dt = my.DocDL(query);
                            dgv_ct.DataSource = dt;
                            dgv_ct.Columns[0].HeaderText = "Mã cuộc thi";
                            dgv_ct.Columns[1].HeaderText = "Tên cuộc thi";
                            dgv_ct.Columns[1].Width = 300;
                            dgv_ct.Columns[2].HeaderText = "Lĩnh vực";
                            dgv_ct.Columns[3].HeaderText = "Năm tổ chức";
                            dgv_ct.Columns[4].HeaderText = "Kinh phí";
                            dgv_ct.Columns[5].HeaderText = "Loại cuộc thi";
                            dgv_ct.Columns[6].HeaderText = "Cấp cuộc thi";


                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm theo mã cuộc thi  !", "Thông báo");
                        }
                    }
                    else if (cbo_tk.SelectedIndex == 1)
                    {
                        try
                        {

                            string query = " select * from CuocThiSTKN where CapCT = N'Cấp Quốc Gia' or  CapCT = N'Cấp Bộ' and TenCuocThi like N'%" + txt_timkiem.Text + "%' ";
                            DataTable dt = my.DocDL(query);
                            dgv_ct.DataSource = dt;
                            dgv_ct.Columns[0].HeaderText = "Mã cuộc thi";
                            dgv_ct.Columns[1].HeaderText = "Tên cuộc thi";
                            dgv_ct.Columns[1].Width = 300;
                            dgv_ct.Columns[2].HeaderText = "Lĩnh vực";
                            dgv_ct.Columns[3].HeaderText = "Năm tổ chức";
                            dgv_ct.Columns[4].HeaderText = "Kinh phí";
                            dgv_ct.Columns[5].HeaderText = "Loại cuộc thi";
                            dgv_ct.Columns[6].HeaderText = "Cấp cuộc thi";

                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm theo tên cuộc thi !", "Thông báo");
                        }
                    }


                }
            }
        }
    }
}
