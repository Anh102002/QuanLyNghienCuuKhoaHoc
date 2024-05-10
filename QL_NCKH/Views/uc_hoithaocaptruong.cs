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
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace QL_NCKH
{
    public partial class uc_hoithaocaptruong : DevExpress.XtraEditors.XtraUserControl
    {
        public uc_hoithaocaptruong()
        {
            InitializeComponent();
        }
        MyClass my = new MyClass();
        Model.XoaHT xoa = new Model.XoaHT();
        private List<string> productListGV;
        private List<string> productListKM;
        private string maht;
        public string Maht
        {
            get { return this.maht; }
            set { this.maht = value; }
        }
        public void loadDLHT()
        {
            try
            {

                string query = " select MaHT,TenHoiThao,NgayToChuc,DiaDiem from HoiThao where CapHoiThao = N'Cấp Trường' ";
                DataTable dt = my.DocDL(query);
                dgv_ht.DataSource = dt;
                dgv_ht.Columns[0].HeaderText = "Mã hội thảo";
                dgv_ht.Columns[1].HeaderText = "Tên hội thảo";
                dgv_ht.Columns[1].Width = 300;
                dgv_ht.Columns[2].HeaderText = "Ngày tổ chức";
                dgv_ht.Columns[3].HeaderText = "Địa điểm";

            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu hội thảo ", "Lỗi");
            }
        }
        public void loadDLNT(string ma)
        {
            try
            {

                string query = @" select TDHT_PhiaNhaTruong.MaGV,GiangVien.HoTen,TDHT_PhiaNhaTruong.ChucVu 
                                   from TDHT_PhiaNhaTruong
                                   Left join GiangVien on TDHT_PhiaNhaTruong.MaGV = GiangVien.MaGV where TDHT_PhiaNhaTruong.MaHT = '" + ma + "' ";
                DataTable dt = my.DocDL(query);
                dgv_nt.DataSource = dt;
                dgv_nt.Columns[0].HeaderText = "Mã giảng viên";
                dgv_nt.Columns[1].HeaderText = "Tên giảng viên";
                dgv_nt.Columns[2].HeaderText = "Chức vụ";
                dgv_nt.Columns[2].Width = 300;


            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu về phía nhà trường ", "Lỗi");
            }
        }

        public void loadDLKM(string ma)
        {
            try
            {

                string query = @" select TDHT_KhachMoi.MaKM,TVNgoaiTruong.HoTen,TDHT_KhachMoi.ChucVu 
                                   from TDHT_KhachMoi
                                   Left join TVNgoaiTruong on TDHT_KhachMoi.MaKM = TVNgoaiTruong.MaKM where TDHT_KhachMoi.MaHT = '" + ma + "' ";
                DataTable dt = my.DocDL(query);
                dgv_km.DataSource = dt;
                dgv_km.Columns[0].HeaderText = "Mã khách mời";
                dgv_km.Columns[1].HeaderText = "Tên khách mời";
                dgv_km.Columns[2].HeaderText = "Chức vụ";
                dgv_km.Columns[2].Width = 300;


            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu phía khách mời ", "Lỗi");
            }
        }
        
        public void loadDLCG(string ma)
        {
            try
            {

                string query = @" select MaCG,TenChuyenGia,HocHam,HocVi,ChucVu from TDHT_ChuyenGia where MaHT = '" + ma + "' ";
                DataTable dt = my.DocDL(query);
                dgv_cg.DataSource = dt;
                dgv_cg.Columns[0].HeaderText = "Mã chuyên gia";
                dgv_cg.Columns[1].HeaderText = "Tên chuyên gia";
                dgv_cg.Columns[2].HeaderText = "Học hàm";
                dgv_cg.Columns[3].HeaderText = "Học vị";
                dgv_cg.Columns[4].HeaderText = "Chức vụ";



            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu chuyên gia ", "Lỗi");
            }
        }
        public void loadDLBTC(string ma)
        {
            try
            {

                string query = @" select BanToChucHT.MaGV,GiangVien.HoTen,BanToChucHT.ChucVu,BanToChucHT.VaiTro from BanToChucHT,GiangVien
                                        WHERE BanToChucHT.MaGV = GiangVien.MaGV and BanToChucHT.MaHT = '" + ma + "'  ";
                DataTable dt = my.DocDL(query);
                dgv_btc.DataSource = dt;
                dgv_btc.Columns[0].HeaderText = "Mã giảng viên";
                dgv_btc.Columns[1].HeaderText = "Tên giảng viên";
                dgv_btc.Columns[1].Width = 150;
                dgv_btc.Columns[2].HeaderText = "Chức vụ";
                dgv_btc.Columns[3].HeaderText = "Vai trò";





            }
            catch (Exception ex)
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu ban tổ chức hội thi {" + ex.Message + "}", "Lỗi");
            }
        }
        private void uc_hoithaocaptruong_Load(object sender, EventArgs e)
        {
            loadDLHT();
            LoadProductListGV();
            LoadProductListKM();
        }
        public bool Ktra()
        {
            if (string.IsNullOrWhiteSpace(txt_maht.Text) || string.IsNullOrWhiteSpace(dtp_ngaytl.Text)
                || string.IsNullOrWhiteSpace(txt_tenht.Text) || string.IsNullOrWhiteSpace(txt_diadiem.Text))
                return false;

            return true;
        }
        public bool KtraMaHT(string ma)
        {
            try
            {
                string sql = "select * from HoiThao where MaHT = '" + ma + "'";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }


            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã hội thảo !", "Thông báo");
            }
            return true;
        }
        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Ktra())
            {
                try
                {
                    if (KtraMaHT(txt_maht.Text))
                    {
                        string ngay = dtp_ngaytl.Value.ToString("yyyy/MM/dd");
                        string sql = "insert into HoiThao values (@Ma,@Ten,@Ngay,@Diadiem,@Cap) ";
                        SqlCommand command = my.SqlCommand(sql);
                        command.Parameters.AddWithValue("@Ma", txt_maht.Text);
                        command.Parameters.AddWithValue("@Ten", txt_tenht.Text);
                        command.Parameters.AddWithValue("@Ngay", ngay);
                        command.Parameters.AddWithValue("@Diadiem", txt_diadiem.Text);
                        command.Parameters.AddWithValue("@Cap", "Cấp Trường");

                        int up = command.ExecuteNonQuery();
                        if (up > 0)
                        {
                            MessageBox.Show("Thêm thông tin thành công", "Thông báo");
                            txt_maht.Clear();
                            txt_tenht.Clear();
                            txt_diadiem.Clear();



                            loadDLHT();

                        }
                    }
                    else
                    {
                        MessageBox.Show("Mã hội thảo này đã tồn tại !", "Thông báo");
                    }



                }
                catch
                {
                    MessageBox.Show("Lỗi ! không thêm thành công ", "Lỗi");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
        }

        private void btn_sua_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Ktra())
            {
                try
                {
                    if (!KtraMaHT(txt_maht.Text))
                    {
                        string ngay = dtp_ngaytl.Value.ToString("yyyy/MM/dd");
                        string sql = "update HoiThao set TenHoiThao=@Ten,NgayToChuc=@Ngay,DiaDiem=@Diadiem,CapHoiThao=@Cap where MaHT=@Ma ";
                        SqlCommand command = my.SqlCommand(sql);
                        command.Parameters.AddWithValue("@Ma", txt_maht.Text);
                        command.Parameters.AddWithValue("@Ten", txt_tenht.Text);
                        command.Parameters.AddWithValue("@Ngay", ngay);
                        command.Parameters.AddWithValue("@Diadiem", txt_diadiem.Text);
                        command.Parameters.AddWithValue("@Cap", "Cấp Trường");

                        int up = command.ExecuteNonQuery();
                        if (up > 0)
                        {
                            MessageBox.Show("Sửa thông tin thành công", "Thông báo");
                            txt_maht.Clear();
                            txt_tenht.Clear();
                            txt_diadiem.Clear();



                            loadDLHT();

                        }
                    }
                    else
                    {
                        MessageBox.Show("Mã hội thảo này không tồn tại !", "Thông báo");
                    }



                }
                catch
                {
                    MessageBox.Show("Lỗi ! không sửa thành công ", "Lỗi");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void btn_xoa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Ktra())
            {
                try
                {
                    if (!KtraMaHT(txt_maht.Text))
                    {


                        DialogResult tb = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                        if (tb == DialogResult.OK)
                        {
                            bool XOABTC = xoa.XoaBTC(txt_maht.Text);
                            //bool XOABTCNT = xoa.XoaBTCNT(txt_maht.Text);
                            bool XOABB = xoa.XoaBBHT(txt_maht.Text);
                            bool XOANT = xoa.XoaNT(txt_maht.Text);
                            bool XOAKM = xoa.XoaKM(txt_maht.Text);
                            bool XOACG = xoa.XoaCG(txt_maht.Text);
                            //bool XOADD = xoa.XoaDD(txt_maht.Text);
                            bool XOAHT = xoa.XoaHoiThao(txt_maht.Text);


                            if (XOABB)
                            {
                                if (XOABTC)
                                {

                                    if (XOANT)
                                    {
                                        if (XOAKM)
                                        {
                                            if (XOACG)
                                            {

                                                if (XOAHT)
                                                {
                                                    MessageBox.Show("Xóa thành công", "Thông báo");
                                                    RESET();
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
                        MessageBox.Show("Mã hội thảo này không tồn tại !", "Thông báo");
                    }



                }
                catch
                {
                    MessageBox.Show("Lỗi ! không xóa thành công ", "Lỗi");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
        }
        public void RESET()
        {
            txt_maht.Clear();
            txt_tenht.Clear();
            txt_diadiem.Clear();
            loadDLHT();

            txt_timkiem.Clear();
            cbo_loai.SelectedIndex = -1;

            txt_chucvugvtv.Clear();
            txt_tengvtv.Clear();
            txt_diadiem.Clear();
            dgv_nt.DataSource = null;
            txt_magvtv.Clear();

            txt_tenkm.Clear();
            txt_makm.Clear();
            txt_chucvukm.Clear();
            dgv_km.DataSource = null;

            txt_tengv.Clear();
            txt_magv.Clear();
            txt_chucvugv.Clear();
            cbo_vaitrogv.SelectedIndex = -1;
            dgv_btc.DataSource = null;

            txt_macg.Clear();
            txt_tencg.Clear();
            txt_hocham.Clear();
            txt_hocvi.Clear();
            txt_chucvucg.Clear();
            dgv_cg.DataSource = null;



        }
        private void btn_refresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            RESET();
        }

        private void barButtonItem7_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            frm_baithamluanHT frm = new frm_baithamluanHT();
            frm.Maht = txt_maht.Text;
            frm.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(cbo_loai.Text))
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

                    if (cbo_loai.SelectedIndex == 0)
                    {
                        try
                        {
                            string query = " select MaHT,TenHoiThao,NgayToChuc,DiaDiem from HoiThao where CapHoiThao = N'Cấp Trường' and MaHT like '%" + txt_timkiem.Text + "%' ";
                            DataTable dt = my.DocDL(query);
                            dgv_ht.DataSource = dt;
                            dgv_ht.Columns[0].HeaderText = "Mã hội thảo";
                            dgv_ht.Columns[1].HeaderText = "Tên hội thảo";
                            dgv_ht.Columns[1].Width = 300;
                            dgv_ht.Columns[2].HeaderText = "Ngày tổ chức";
                            dgv_ht.Columns[3].HeaderText = "Địa điểm";



                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm theo mã hội thảo  !", "Thông báo");
                        }
                    }
                    else if (cbo_loai.SelectedIndex == 1)
                    {
                        try
                        {

                            string query = " select MaHT,TenHoiThao,NgayToChuc,DiaDiem from HoiThao where CapHoiThao = N'Cấp Trường' and TenHoiThao like N'%" + txt_timkiem.Text + "%' ";
                            DataTable dt = my.DocDL(query);
                            dgv_ht.DataSource = dt;
                            dgv_ht.Columns[0].HeaderText = "Mã hội thảo";
                            dgv_ht.Columns[1].HeaderText = "Tên hội thảo";
                            dgv_ht.Columns[1].Width = 300;
                            dgv_ht.Columns[2].HeaderText = "Ngày tổ chức";
                            dgv_ht.Columns[3].HeaderText = "Địa điểm";


                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm theo tên hội thảo !", "Thông báo");
                        }
                    }


                }
            }
        }

        private void dgv_ht_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_maht.Text = dgv_ht.CurrentRow.Cells[0].Value.ToString();
                txt_tenht.Text = dgv_ht.CurrentRow.Cells[1].Value.ToString();
                dtp_ngaytl.Text = dgv_ht.CurrentRow.Cells[2].Value.ToString();
                txt_diadiem.Text = dgv_ht.CurrentRow.Cells[3].Value.ToString();
                Maht = txt_maht.Text;
                string ma = Maht;
                loadDLNT(ma);
                loadDLKM(ma);
                loadDLCG(ma);
                loadDLBTC(ma);
            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị thông tin hội thảo", "Lỗi");
            }
        }

        private void dgv_nt_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_magvtv.Text = dgv_nt.CurrentRow.Cells[0].Value.ToString();
                txt_tengvtv.Text = dgv_nt.CurrentRow.Cells[1].Value.ToString();
                txt_chucvugvtv.Text = dgv_nt.CurrentRow.Cells[2].Value.ToString();


            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị thông tin nhà trường", "Lỗi");
            }
        }

        private void dgv_km_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_makm.Text = dgv_km.CurrentRow.Cells[0].Value.ToString();
                txt_tenkm.Text = dgv_km.CurrentRow.Cells[1].Value.ToString();
                txt_chucvukm.Text = dgv_km.CurrentRow.Cells[2].Value.ToString();


            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị thông tin khách mời", "Lỗi");
            }
        }

        private void dgv_cg_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_macg.Text = dgv_cg.CurrentRow.Cells[0].Value.ToString();
                txt_tencg.Text = dgv_cg.CurrentRow.Cells[1].Value.ToString();
                txt_hocham.Text = dgv_cg.CurrentRow.Cells[2].Value.ToString();
                txt_hocvi.Text = dgv_cg.CurrentRow.Cells[3].Value.ToString();
                txt_chucvucg.Text = dgv_cg.CurrentRow.Cells[4].Value.ToString();


            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị thông tin chuyên gia", "Lỗi");
            }
        }

        private void dgv_btc_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_magv.Text = dgv_btc.CurrentRow.Cells[0].Value.ToString();
                txt_tengv.Text = dgv_btc.CurrentRow.Cells[1].Value.ToString();
                txt_chucvugv.Text = dgv_btc.CurrentRow.Cells[2].Value.ToString();
                cbo_vaitrogv.Text = dgv_btc.CurrentRow.Cells[3].Value.ToString();

            }
            catch
            {
                MessageBox.Show("Lỗi hiển thị dữ liệu ban tổ chức  ", "Lỗi");
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
        private void ShowSuggestions(List<string> suggestions)
        {
            list_nt.Items.Clear();
            list_nt.Items.AddRange(suggestions.ToArray());

            list_nt.Visible = suggestions.Any();
        }
        private void txt_magvtv_TextChanged(object sender, EventArgs e)
        {
            string searchTerm = txt_magvtv.Text.ToLower();
            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                List<string> filteredProducts = productListGV
               .Where(product => product.ToLower().Contains(searchTerm))
               .ToList();

                if (filteredProducts != null)
                {
                    ShowSuggestions(filteredProducts);
                }
                else
                {
                    list_nt.Visible = false;

                }


            }
            else
            {
                list_nt.Visible = false;
                txt_tengvtv.Clear();
            }
        }

        private void list_nt_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_nt.SelectedItem != null)
            {
                string selectedProduct = list_nt.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_magvtv.Text = selectedProduct;
                    list_nt.Visible = false;
                    string sql = "select HoTen from GiangVien where MaGV = '" + selectedProduct + "' ";
                    DataTable tb = my.DocDL(sql);
                    if (tb.Rows.Count > 0)
                    {
                        string hoten = tb.Rows[0][0].ToString();
                        txt_tengvtv.Text = hoten;
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
        private void ShowSuggestionsBTC(List<string> suggestions)
        {
            list_btc.Items.Clear();
            list_btc.Items.AddRange(suggestions.ToArray());

            list_btc.Visible = suggestions.Any();
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
                    ShowSuggestionsBTC(filteredProducts);
                }
                else
                {
                    list_btc.Visible = false;

                }


            }
            else
            {
                list_btc.Visible = false;
                txt_tengv.Clear();
            }
        }

        private void list_btc_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_btc.SelectedItem != null)
            {
                string selectedProduct = list_btc.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_magv.Text = selectedProduct;
                    list_btc.Visible = false;
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
        public bool KtraMaNT(string ma, string maHT)
        {
            try
            {
                string sql = "select * from TDHT_PhiaNhaTruong where MaGV = '" + ma + "'  and MaHT='" + maHT + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }


            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã phía nhà trường !", "Lỗi");
            }
            return true;
        }
        private void btn_joingvtv_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_magvtv.Text) || string.IsNullOrWhiteSpace(txt_tengvtv.Text) || string.IsNullOrWhiteSpace(txt_chucvugvtv.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {

                    if (string.IsNullOrWhiteSpace(Maht) || dgv_nt.DataSource == null)
                    {
                        MessageBox.Show("Vui lòng chọn hội thảo !", "Thông báo");
                    }
                    else
                    {
                        string maht = Maht;
                        if (KtraMaNT(txt_magvtv.Text, maht))
                        {


                            string sql = "insert into TDHT_PhiaNhaTruong values (@Magv,@Maht,@Chucvu)";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Magv", txt_magvtv.Text);
                            comd.Parameters.AddWithValue("@Maht", maht);
                            comd.Parameters.AddWithValue("@Chucvu", txt_chucvugvtv.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Thêm thành công !", "Thông báo");
                                txt_tengvtv.Clear();
                                txt_magvtv.Clear();
                                txt_chucvugvtv.Clear();
                                loadDLNT(maht);
                            }
                            else
                            {
                                MessageBox.Show("Thêm không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("Giảng viên này đã tham dự !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi thêm phía nhà trường!", "Thông báo");
                }
            }
        }

        private void btn_cancelgvtv_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_magvtv.Text) || string.IsNullOrWhiteSpace(txt_tengvtv.Text) || string.IsNullOrWhiteSpace(txt_chucvugvtv.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {

                    if (string.IsNullOrWhiteSpace(Maht) || dgv_nt.DataSource == null)
                    {
                        MessageBox.Show("Vui lòng chọn hội thảo !", "Thông báo");
                    }
                    else
                    {
                        string maht = Maht;
                        if (!KtraMaNT(txt_magvtv.Text, maht))
                        {


                            string sql = "delete from TDHT_PhiaNhaTruong where  MaGV = @Magv and MaHT=@Maht ";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Magv", txt_magvtv.Text);
                            comd.Parameters.AddWithValue("@Maht", maht);
                            //comd.Parameters.AddWithValue("@Chucvu", txt_chucvugvtv.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa thành công !", "Thông báo");
                                txt_tengvtv.Clear();
                                txt_magvtv.Clear();
                                txt_chucvugvtv.Clear();
                                loadDLNT(maht);
                            }
                            else
                            {
                                MessageBox.Show("Xóa không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("Giảng viên này không có tham dự !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi xóa phía nhà trường!", "Thông báo");
                }
            }
        }

        private void btn_suagvtv_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_magvtv.Text) || string.IsNullOrWhiteSpace(txt_tengvtv.Text) || string.IsNullOrWhiteSpace(txt_chucvugvtv.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {

                    if (string.IsNullOrWhiteSpace(Maht) || dgv_nt.DataSource == null)
                    {
                        MessageBox.Show("Vui lòng chọn hội thảo !", "Thông báo");
                    }
                    else
                    {
                        string maht = Maht;
                        if (!KtraMaNT(txt_magvtv.Text, maht))
                        {


                            string sql = "update TDHT_PhiaNhaTruong set ChucVu=@Chucvu where  MaGV=@Magv and MaHT=@Maht ";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Magv", txt_magvtv.Text);
                            comd.Parameters.AddWithValue("@Maht", maht);
                            comd.Parameters.AddWithValue("@Chucvu", txt_chucvugvtv.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa thành công !", "Thông báo");
                                txt_tengvtv.Clear();
                                txt_magvtv.Clear();
                                txt_chucvugvtv.Clear();
                                loadDLNT(maht);
                            }
                            else
                            {
                                MessageBox.Show("Sửa không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("Giảng viên này không có tham dự !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi sửa phía nhà trường!", "Thông báo");
                }
            }
        }
        public bool KtraMaKM(string ma, string maHT)
        {
            try
            {
                string sql = "select * from TDHT_KhachMoi where MaKM = '" + ma + "'  and MaHT='" + maHT + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }


            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã phía khách mời !", "Lỗi");
            }
            return true;
        }
        private void btn_joinkm_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_makm.Text) || string.IsNullOrWhiteSpace(txt_tenkm.Text) || string.IsNullOrWhiteSpace(txt_chucvukm.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {

                    if (string.IsNullOrWhiteSpace(Maht) || dgv_km.DataSource == null)
                    {
                        MessageBox.Show("Vui lòng chọn hội thảo !", "Thông báo");
                    }
                    else
                    {
                        string maht = Maht;
                        if (KtraMaKM(txt_makm.Text, maht))
                        {


                            string sql = "insert into TDHT_KhachMoi values (@Maht,@Makm,@Chucvu)";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Makm", txt_makm.Text);
                            comd.Parameters.AddWithValue("@Maht", maht);
                            comd.Parameters.AddWithValue("@Chucvu", txt_chucvukm.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Thêm thành công !", "Thông báo");
                                txt_tenkm.Clear();
                                txt_makm.Clear();
                                txt_chucvukm.Clear();
                                loadDLKM(maht);
                            }
                            else
                            {
                                MessageBox.Show("Thêm không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("Khách mời này đã tham dự !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi thêm phía khách mời!", "Thông báo");
                }
            }
        }

        private void btn_cancelkm_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_makm.Text) || string.IsNullOrWhiteSpace(txt_tenkm.Text) || string.IsNullOrWhiteSpace(txt_chucvukm.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {

                    if (string.IsNullOrWhiteSpace(Maht) || dgv_km.DataSource == null)
                    {
                        MessageBox.Show("Vui lòng chọn hội thảo !", "Thông báo");
                    }
                    else
                    {
                        string maht = Maht;
                        if (!KtraMaKM(txt_makm.Text, maht))
                        {


                            string sql = "delete from TDHT_KhachMoi where MaKM = @Makm and MaHT= @Maht ";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Makm", txt_makm.Text);
                            comd.Parameters.AddWithValue("@Maht", maht);
                            //comd.Parameters.AddWithValue("@Chucvu", txt_chucvugvtv.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa thành công !", "Thông báo");
                                txt_tenkm.Clear();
                                txt_makm.Clear();
                                txt_chucvukm.Clear();
                                loadDLKM(maht);
                            }
                            else
                            {
                                MessageBox.Show("Xóa không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("Khách mời này không có tham dự !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi xóa phía khách mời!", "Thông báo");
                }
            }
        }

        private void btn_suakm_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_makm.Text) || string.IsNullOrWhiteSpace(txt_tenkm.Text) || string.IsNullOrWhiteSpace(txt_chucvukm.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {

                    if (string.IsNullOrWhiteSpace(Maht) || dgv_km.DataSource == null)
                    {
                        MessageBox.Show("Vui lòng chọn hội thảo !", "Thông báo");
                    }
                    else
                    {
                        string maht = Maht;
                        if (!KtraMaKM(txt_makm.Text, maht))
                        {


                            string sql = "update TDHT_KhachMoi set ChucVu = @Chucvu where MaKM= @Makm and MaHT= @Maht";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Makm", txt_makm.Text);
                            comd.Parameters.AddWithValue("@Maht", maht);
                            comd.Parameters.AddWithValue("@Chucvu", txt_chucvukm.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa thành công !", "Thông báo");
                                txt_tenkm.Clear();
                                txt_makm.Clear();
                                txt_chucvukm.Clear();
                                loadDLKM(maht);
                            }
                            else
                            {
                                MessageBox.Show("Sửa không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("Khách mời này không có tham dự !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi sửa phía khách mời!", "Thông báo");
                }
            }
        }
        public bool KtraMaCG(string ma)
        {
            try
            {
                string sql = "select * from TDHT_ChuyenGia where MaCG = '" + ma + "'  ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }


            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã chuyên gia !", "Lỗi");
            }
            return true;
        }
        private void btn_joincg_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_macg.Text) || string.IsNullOrWhiteSpace(txt_tencg.Text)
                || string.IsNullOrWhiteSpace(txt_hocham.Text) || string.IsNullOrWhiteSpace(txt_hocvi.Text) || string.IsNullOrWhiteSpace(txt_chucvucg.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {

                    if (string.IsNullOrWhiteSpace(Maht) || dgv_cg.DataSource == null)
                    {
                        MessageBox.Show("Vui lòng chọn hội thảo !", "Thông báo");
                    }
                    else
                    {
                        string maht = Maht;
                        if (KtraMaCG(txt_macg.Text))
                        {


                            string sql = "insert into TDHT_ChuyenGia values (@Macg,@Maht,@Tencg,@Hocham,@Hocvi,@Chucvu)";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Macg", txt_macg.Text);
                            comd.Parameters.AddWithValue("@Maht", maht);
                            comd.Parameters.AddWithValue("@Tencg", txt_tencg.Text);
                            comd.Parameters.AddWithValue("@Hocham", txt_hocham.Text);
                            comd.Parameters.AddWithValue("@Hocvi", txt_hocvi.Text);
                            comd.Parameters.AddWithValue("@Chucvu", txt_chucvucg.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Thêm thành công !", "Thông báo");
                                txt_macg.Clear();
                                txt_tencg.Clear();
                                txt_chucvucg.Clear();
                                txt_hocham.Clear();
                                txt_hocvi.Clear();
                                loadDLCG(maht);
                            }
                            else
                            {
                                MessageBox.Show("Thêm không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("Chuyên gia này đã tham dự !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi thêm chuyên gia!", "Thông báo");
                }
            }
        }

        private void btn_cancelcg_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_macg.Text) || string.IsNullOrWhiteSpace(txt_tencg.Text)
                || string.IsNullOrWhiteSpace(txt_hocham.Text) || string.IsNullOrWhiteSpace(txt_hocvi.Text) || string.IsNullOrWhiteSpace(txt_chucvucg.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {

                    if (string.IsNullOrWhiteSpace(Maht) || dgv_cg.DataSource == null)
                    {
                        MessageBox.Show("Vui lòng chọn hội thảo !", "Thông báo");
                    }
                    else
                    {
                        string maht = Maht;
                        if (!KtraMaCG(txt_macg.Text))
                        {


                            string sql = "delete from TDHT_ChuyenGia where MaCG=@Macg ";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Macg", txt_macg.Text);
                            //comd.Parameters.AddWithValue("@Maht", maht);
                            //comd.Parameters.AddWithValue("@Tencg", txt_tencg.Text);
                            //comd.Parameters.AddWithValue("@Hocham", txt_hocham.Text);
                            //comd.Parameters.AddWithValue("@Hocvi", txt_hocvi.Text);
                            //comd.Parameters.AddWithValue("@Chucvu", txt_chucvucg.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa thành công !", "Thông báo");
                                txt_macg.Clear();
                                txt_tencg.Clear();
                                txt_chucvucg.Clear();
                                txt_hocham.Clear();
                                txt_hocvi.Clear();
                                loadDLCG(maht);
                            }
                            else
                            {
                                MessageBox.Show("Xóa không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("Chuyên gia này không có tham dự !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi xóa chuyên gia!", "Thông báo");
                }
            }
        }

        private void btn_suacg_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_macg.Text) || string.IsNullOrWhiteSpace(txt_tencg.Text)
                || string.IsNullOrWhiteSpace(txt_hocham.Text) || string.IsNullOrWhiteSpace(txt_hocvi.Text) || string.IsNullOrWhiteSpace(txt_chucvucg.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {

                    if (string.IsNullOrWhiteSpace(Maht) || dgv_cg.DataSource == null)
                    {
                        MessageBox.Show("Vui lòng chọn hội thảo !", "Thông báo");
                    }
                    else
                    {
                        string maht = Maht;
                        if (!KtraMaCG(txt_macg.Text))
                        {


                            string sql = "update TDHT_ChuyenGia  set  TenChuyenGia=@Tencg,HocHam=@Hocham,HocVi=@Hocvi,ChucVu=@Chucvu where MaCG=@Macg and MaHT=@Maht ";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Macg", txt_macg.Text);
                            comd.Parameters.AddWithValue("@Maht", maht);
                            comd.Parameters.AddWithValue("@Tencg", txt_tencg.Text);
                            comd.Parameters.AddWithValue("@Hocham", txt_hocham.Text);
                            comd.Parameters.AddWithValue("@Hocvi", txt_hocvi.Text);
                            comd.Parameters.AddWithValue("@Chucvu", txt_chucvucg.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa thành công !", "Thông báo");
                                txt_macg.Clear();
                                txt_tencg.Clear();
                                txt_chucvucg.Clear();
                                txt_hocham.Clear();
                                txt_hocvi.Clear();
                                loadDLCG(maht);
                            }
                            else
                            {
                                MessageBox.Show("Sửa không thành công !", "Thông báo");
                            }





                        }
                        else
                        {
                            MessageBox.Show("Chuyên gia này không có tham dự !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi sửa chuyên gia!", "Thông báo");
                }
            }
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
        private void btn_joinld_Click(object sender, EventArgs e)
        {
            string ma = Maht;
            if (string.IsNullOrWhiteSpace(ma) || dgv_btc.DataSource == null)
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
        public void ExcelExportDSHT()
        {
            try
            {
                Excel.Application oExcel = new Excel.Application();
                Excel.Workbook oBook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

                Excel.Worksheet oSheet = (Excel.Worksheet)oBook.Worksheets[1];

                Excel.Range head = oSheet.get_Range("A1", "D1");

                head.MergeCells = true;

                head.Value2 = "DANH SÁCH HỘI THẢO CẤP TRƯỜNG";

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

                //Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                //cl5.Value = "Kinh phí";

                //Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                //cl6.Value = "Loại cuộc thi";





                Excel.Range rowHead = oSheet.get_Range("A3", "D3");
                rowHead.Font.Bold = true;
                rowHead.Font.Size = 13;
                rowHead.Font.Name = "Times New Roman";
                rowHead.Borders.LineStyle = Excel.Constants.xlSolid;
                rowHead.Interior.ColorIndex = 6;
                rowHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int line = 4;
                for (int i = 0; i < dgv_ht.Rows.Count - 1; i++)
                {
                    Excel.Range line1 = oSheet.get_Range("A" + (line + i).ToString(), "A" + (line + i).ToString());
                    line1.Value = dgv_ht.Rows[i].Cells[0].Value.ToString();
                    line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line1.Borders.LineStyle = Excel.Constants.xlSolid;
                    line1.Font.Name = "Times New Roman";

                    Excel.Range line2 = oSheet.get_Range("B" + (line + i).ToString(), "B" + (line + i).ToString());
                    line2.Value = dgv_ht.Rows[i].Cells[1].Value.ToString();
                    line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line2.Borders.LineStyle = Excel.Constants.xlSolid;
                    line2.Font.Name = "Times New Roman";

                    Excel.Range line3 = oSheet.get_Range("C" + (line + i).ToString(), "C" + (line + i).ToString());
                    line3.Value = dgv_ht.Rows[i].Cells[2].Value.ToString();
                    line3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line3.Borders.LineStyle = Excel.Constants.xlSolid;
                    line3.Font.Name = "Times New Roman";

                    Excel.Range line4 = oSheet.get_Range("D" + (line + i).ToString(), "D" + (line + i).ToString());
                    line4.Value = dgv_ht.Rows[i].Cells[3].Value.ToString();
                    line4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line4.Borders.LineStyle = Excel.Constants.xlSolid;
                    line4.Font.Name = "Times New Roman";


                    //Excel.Range line5 = oSheet.get_Range("E" + (line + i).ToString(), "E" + (line + i).ToString());
                    //line5.Value = dgv_ht.Rows[i].Cells[4].Value.ToString();
                    //line5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    //line5.Borders.LineStyle = Excel.Constants.xlSolid;
                    //line5.Font.Name = "Times New Roman";

                    //Excel.Range line6 = oSheet.get_Range("F" + (line + i).ToString(), "F" + (line + i).ToString());
                    //line6.Value = dgv_ht.Rows[i].Cells[5].Value.ToString();
                    //line6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    //line6.Borders.LineStyle = Excel.Constants.xlSolid;
                    //line6.Font.Name = "Times New Roman";



                }


                oSheet.Name = "DSHTCT";
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
            ExcelExportDSHT();
        }


        public DataTable LayDuLieuBaoCao()
        {

            string query = " select MaHT,TenHoiThao,NgayToChuc,DiaDiem from HoiThao where CapHoiThao = N'Cấp Trường' ";
            DataTable dataTable = my.DocDL(query);

            return dataTable;
        }
        public void excelCTHT()
        {
            try
            {

                DataTable dataTable = LayDuLieuBaoCao();


                Excel.Application oExcel = new Excel.Application();
                Excel.Workbook workbook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet oSheet = (Excel.Worksheet)workbook.Worksheets[1];




                Excel.Range head = oSheet.get_Range("A1", "H1");

                head.MergeCells = true;

                head.Value2 = "DANH SÁCH CHI TIẾT HỘI THẢO CẤP TRƯỜNG  ";

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
                cl5.Value = "Tham dự hội thảo phía nhà trường";

                Excel.Range cl10 = oSheet.get_Range("F3", "F3");
                cl10.Value = "Tham dự hội thảo phía khách mời";

                Excel.Range cl6 = oSheet.get_Range("G3", "G3");
                cl6.Value = "Chuyên gia tham dự hội thảo";

                Excel.Range cl7 = oSheet.get_Range("H3", "H3");
                cl7.Value = "Ban tổ chức hội thảo";

                //Excel.Range cl8 = oSheet.get_Range("I3", "I3");
                //cl8.Value = "";

                //Excel.Range cl9 = oSheet.get_Range("J3", "J3");
                //cl9.Value = "Giảng viên hướng dẫn";





                Excel.Range rowHead = oSheet.get_Range("A3", "H3");
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
                    string nt = @" select TDHT_PhiaNhaTruong.MaGV,GiangVien.HoTen,TDHT_PhiaNhaTruong.ChucVu 
                                   from TDHT_PhiaNhaTruong ,GiangVien where TDHT_PhiaNhaTruong.MaGV = GiangVien.MaGV and TDHT_PhiaNhaTruong.MaHT = '" + maCT + "' ";

                    DataTable dt = my.DocDL(nt);

                    Excel.Range line1 = oSheet.get_Range("E" + (lines).ToString(), "E" + (lines).ToString());
                    Excel.Range line2 = oSheet.get_Range("F" + (lines).ToString(), "F" + (lines).ToString());
                    Excel.Range line3 = oSheet.get_Range("G" + (lines).ToString(), "G" + (lines).ToString());
                    Excel.Range line4 = oSheet.get_Range("H" + (lines).ToString(), "H" + (lines).ToString());
                    //Excel.Range line5 = oSheet.get_Range("I" + (lines).ToString(), "I" + (lines).ToString());

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        string maDT = dt.Rows[row][0].ToString();

                        string cel = dt.Rows[row]["MaGV"].ToString() + "-" + dt.Rows[row]["HoTen"].ToString() + "-" + dt.Rows[row]["ChucVu"].ToString() + "\n";
                        line1.Value += cel;
                        //

                    }
                    line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line1.Borders.LineStyle = Excel.Constants.xlSolid;
                    line1.Font.Name = "Times New Roman";
                    //

                    //
                    string km = @" select TDHT_KhachMoi.MaKM,TVNgoaiTruong.HoTen,TDHT_KhachMoi.ChucVu 
                                   from TDHT_KhachMoi
                                   Left join TVNgoaiTruong on TDHT_KhachMoi.MaKM = TVNgoaiTruong.MaKM where TDHT_KhachMoi.MaHT = '" + maCT + "' ";

                    DataTable dt_bgk = my.DocDL(km);



                    for (int r = 0; r < dt_bgk.Rows.Count; r++)
                    {

                        string celSV = dt_bgk.Rows[r]["MaKM"].ToString() + "-" + dt_bgk.Rows[r]["HoTen"].ToString() + "-" + dt_bgk.Rows[r]["ChucVu"].ToString() + "\n";
                        line2.Value += celSV;


                    }
                    line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line2.Borders.LineStyle = Excel.Constants.xlSolid;
                    line2.Font.Name = "Times New Roman";

                    //
                    



                    //
                    //
                    string cg = @" select MaCG,TenChuyenGia,ChucVu from TDHT_ChuyenGia where MaHT = '" + maCT + "' ";

                    DataTable dt_bgknt = my.DocDL(cg);



                    for (int r1 = 0; r1 < dt_bgknt.Rows.Count; r1++)
                    {

                        string celSVNT = dt_bgknt.Rows[r1]["MaCG"].ToString() + "-" + dt_bgknt.Rows[r1]["TenChuyenGia"].ToString() + "-" + dt_bgknt.Rows[r1]["ChucVu"].ToString() + "\n";
                        line3.Value += celSVNT;


                    }
                    line3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line3.Borders.LineStyle = Excel.Constants.xlSolid;

                    line3.Font.Name = "Times New Roman";

                    //
                    //
                    //

                    string dd = @" select BanToChucHT.MaGV,GiangVien.HoTen,BanToChucHT.ChucVu,BanToChucHT.VaiTro from BanToChucHT,GiangVien
                                        WHERE BanToChucHT.MaGV = GiangVien.MaGV and BanToChucHT.MaHT = '" + maCT + "'  ";

                    DataTable dt_dd = my.DocDL(dd);



                    for (int r2 = 0; r2 < dt_dd.Rows.Count; r2++)
                    {

                        string celDD = dt_dd.Rows[r2]["MaGV"].ToString() + "-" + dt_dd.Rows[r2]["HoTen"].ToString() + "-" + dt_dd.Rows[r2]["ChucVu"].ToString() + "-" + dt_dd.Rows[r2]["VaiTro"].ToString() + "\n";
                        line4.Value += celDD;


                    }
                    line4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line4.Borders.LineStyle = Excel.Constants.xlSolid;

                    line4.Font.Name = "Times New Roman";


                    //
                    //
                    lines++;




                }

                oSheet.Name = "DSCTHTCT";
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
            excelCTHT();
        }
        public DataTable LayDuLieuBaoCao1HT(string ma)
        {

            string query = " select MaHT,TenHoiThao,NgayToChuc,DiaDiem from HoiThao where CapHoiThao = N'Cấp Trường' and MaHT = '"+ma+"' ";
            DataTable dataTable = my.DocDL(query);

            return dataTable;
        }
        public void excel1HT()
        {
            try
            {
                if(Ktra())
                {
                    string ma = txt_maht.Text;
                    DataTable dataTable = LayDuLieuBaoCao1HT(ma);

                
                


                Excel.Application oExcel = new Excel.Application();
                Excel.Workbook workbook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet oSheet = (Excel.Worksheet)workbook.Worksheets[1];




                Excel.Range head = oSheet.get_Range("A1", "H1");

                head.MergeCells = true;

                head.Value2 = "DANH SÁCH CHI TIẾT HỘI THẢO CẤP TRƯỜNG  ";

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
                cl5.Value = "Tham dự hội thảo phía nhà trường";

                Excel.Range cl10 = oSheet.get_Range("F3", "F3");
                cl10.Value = "Tham dự hội thảo phía khách mời";

                Excel.Range cl6 = oSheet.get_Range("G3", "G3");
                cl6.Value = "Chuyên gia tham dự hội thảo";

                Excel.Range cl7 = oSheet.get_Range("H3", "H3");
                cl7.Value = "Ban tổ chức hội thảo";

                //Excel.Range cl8 = oSheet.get_Range("I3", "I3");
                //cl8.Value = "";

                //Excel.Range cl9 = oSheet.get_Range("J3", "J3");
                //cl9.Value = "Giảng viên hướng dẫn";





                Excel.Range rowHead = oSheet.get_Range("A3", "H3");
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
                    string nt = @" select TDHT_PhiaNhaTruong.MaGV,GiangVien.HoTen,TDHT_PhiaNhaTruong.ChucVu 
                                   from TDHT_PhiaNhaTruong ,GiangVien where TDHT_PhiaNhaTruong.MaGV = GiangVien.MaGV and TDHT_PhiaNhaTruong.MaHT = '" + maCT + "' ";

                    DataTable dt = my.DocDL(nt);

                    Excel.Range line1 = oSheet.get_Range("E" + (lines).ToString(), "E" + (lines).ToString());
                    Excel.Range line2 = oSheet.get_Range("F" + (lines).ToString(), "F" + (lines).ToString());
                    Excel.Range line3 = oSheet.get_Range("G" + (lines).ToString(), "G" + (lines).ToString());
                    Excel.Range line4 = oSheet.get_Range("H" + (lines).ToString(), "H" + (lines).ToString());
                    //Excel.Range line5 = oSheet.get_Range("I" + (lines).ToString(), "I" + (lines).ToString());

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        string maDT = dt.Rows[row][0].ToString();

                        string cel = dt.Rows[row]["MaGV"].ToString() + "-" + dt.Rows[row]["HoTen"].ToString() + "-" + dt.Rows[row]["ChucVu"].ToString() + "\n";
                        line1.Value += cel;
                        //

                    }
                    line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line1.Borders.LineStyle = Excel.Constants.xlSolid;
                    line1.Font.Name = "Times New Roman";
                    //

                    //
                    string km = @" select TDHT_KhachMoi.MaKM,TVNgoaiTruong.HoTen,TDHT_KhachMoi.ChucVu 
                                   from TDHT_KhachMoi
                                   Left join TVNgoaiTruong on TDHT_KhachMoi.MaKM = TVNgoaiTruong.MaKM where TDHT_KhachMoi.MaHT = '" + maCT + "' ";

                    DataTable dt_bgk = my.DocDL(km);



                    for (int r = 0; r < dt_bgk.Rows.Count; r++)
                    {

                        string celSV = dt_bgk.Rows[r]["MaKM"].ToString() + "-" + dt_bgk.Rows[r]["HoTen"].ToString() + "-" + dt_bgk.Rows[r]["ChucVu"].ToString() + "\n";
                        line2.Value += celSV;


                    }
                    line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line2.Borders.LineStyle = Excel.Constants.xlSolid;
                    line2.Font.Name = "Times New Roman";

                    //




                    //
                    //
                    string cg = @" select MaCG,TenChuyenGia,ChucVu from TDHT_ChuyenGia where MaHT = '" + maCT + "' ";

                    DataTable dt_bgknt = my.DocDL(cg);



                    for (int r1 = 0; r1 < dt_bgknt.Rows.Count; r1++)
                    {

                        string celSVNT = dt_bgknt.Rows[r1]["MaCG"].ToString() + "-" + dt_bgknt.Rows[r1]["TenChuyenGia"].ToString() + "-" + dt_bgknt.Rows[r1]["ChucVu"].ToString() + "\n";
                        line3.Value += celSVNT;


                    }
                    line3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line3.Borders.LineStyle = Excel.Constants.xlSolid;

                    line3.Font.Name = "Times New Roman";

                    //
                    //
                    //

                    string dd = @" select BanToChucHT.MaGV,GiangVien.HoTen,BanToChucHT.ChucVu,BanToChucHT.VaiTro from BanToChucHT,GiangVien
                                        WHERE BanToChucHT.MaGV = GiangVien.MaGV and BanToChucHT.MaHT = '" + maCT + "'  ";

                    DataTable dt_dd = my.DocDL(dd);



                    for (int r2 = 0; r2 < dt_dd.Rows.Count; r2++)
                    {

                        string celDD = dt_dd.Rows[r2]["MaGV"].ToString() + "-" + dt_dd.Rows[r2]["HoTen"].ToString() + "-" + dt_dd.Rows[r2]["ChucVu"].ToString() + "-" + dt_dd.Rows[r2]["VaiTro"].ToString() + "\n";
                        line4.Value += celDD;


                    }
                    line4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line4.Borders.LineStyle = Excel.Constants.xlSolid;

                    line4.Font.Name = "Times New Roman";


                    //
                    //
                    lines++;




                }

                oSheet.Name = "DSCTHTCT";
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
                    MessageBox.Show("Vui lòng chọn hội thảo cần export dữ liệu", "Thông báo");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi xuất báo cáo: {ex.Message}");
            }
        }

        private void barButtonItem4_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            excel1HT();
        }
    }
}
