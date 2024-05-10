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
using System.Threading;

namespace QL_NCKH
{
    public partial class uc_detaicapCoSo : DevExpress.XtraEditors.XtraUserControl
    {

        MyClass my = new MyClass();
        private List<string> productList;
        private string madt;

        public uc_detaicapCoSo()
        {
            InitializeComponent();
        }

        public string Madt
        {
            get { return this.madt; }
            set { this.madt = value; }
        }
        public void LoadDL()
        {
            string cap = "Cấp Cơ Sở";
            string query = "select DeTai.MaDeTai,DeTai.TenDeTai,DeTai.Khoa,DeTai.LinhVuc,TienDoDeTai.NgayBatDau,TienDoDeTai.NgayKetThuc,TienDoDeTai.TienDo from DeTai,TienDoDeTai  where CapDeTai = N'" + cap + "' and TienDoDeTai.MaDeTai = DeTai.MaDeTai and DoiTuong = N'Giảng viên' ";
            DataTable dt = my.DocDL(query);
            dgv_dt.DataSource = dt;
            dgv_dt.Columns[0].HeaderText = "Mã đề tài";
            dgv_dt.Columns[1].HeaderText = "Tên đề tài";
            dgv_dt.Columns[2].HeaderText = "Khoa";
            dgv_dt.Columns[3].HeaderText = "Lĩnh vực";
            dgv_dt.Columns[4].HeaderText = "Ngày bắt đầu";
            dgv_dt.Columns[5].HeaderText = "Ngày kết thúc";
            dgv_dt.Columns[6].HeaderText = "Tiến độ";

            cb_khoa.Items.Clear();
            cb_khoa.Items.Add("Công nghệ thông tin");
            cb_khoa.Items.Add("Điện tử");
            cb_khoa.Items.Add("Cơ khí");
            cb_khoa.Items.Add("Kế toán Kiểm toán");
            cb_khoa.Items.Add("Ngoại Ngữ");
            cb_khoa.Items.Add("Du lịch và khách sạn");
            cb_khoa.Items.Add("Tài chình ngân hàng và BH");
            cb_khoa.Items.Add("Thương mại");
            cb_khoa.Items.Add("Điện - Tự động hóa");
            cb_khoa.Items.Add("Diệt may và Thời trang");
            cb_khoa.Items.Add(" Lý luận chính trị và Pháp luật");
            cb_khoa.Items.Add("Quản trị kinh doanh");
            cb_khoa.Items.Add("Quản trị & Marketing");
            cb_khoa.Items.Add("Công nghệ thực phẩm");
            cb_khoa.Items.Add("Khoa học ứng dụng");

        }

        public void LoadDLGV(string ma)
        {
            string query = "SELECT ChiTietGVDeTai.MaGV, GiangVien.HoTen, ChiTietGVDeTai.ChucVu FROM ChiTietGVDeTai, GiangVien WHERE ChiTietGVDeTai.MaGV = GiangVien.MaGV AND MaDeTai = '" + ma + "' ";
            DataTable dt = my.DocDL(query);
            dgv_gv.DataSource = dt;
            dgv_gv.Columns[0].HeaderText = "Mã giảng viên";
            dgv_gv.Columns[1].HeaderText = "Tên giảng viên";
            dgv_gv.Columns[2].HeaderText = "Chức vụ";

        }
        public void LoadDLHD(string ma)
        {
            string query = "SELECT HoiDong.MaGV, GiangVien.HoTen, HoiDong.ChucVuHD FROM HoiDong,GiangVien WHERE HoiDong.MaGV = GiangVien.MaGV AND MaDeTai = '" + ma + "' ";
            DataTable dt = my.DocDL(query);
            dgv_hd.DataSource = dt;
            dgv_hd.Columns[0].HeaderText = "Mã giảng viên";
            dgv_hd.Columns[1].HeaderText = "Tên giảng viên";
            dgv_hd.Columns[2].HeaderText = "Chức vụ";

        }
        private void uc_detaicapCoSo_Load(object sender, EventArgs e)
        {
            try
            {
                LoadDL();
                LoadProductList();
            }
            catch
            {
                MessageBox.Show("Lỗi hiển thị dữ liệu !", "Thông báo");
            }
        }

        public bool Ktra()
        {
            if (string.IsNullOrWhiteSpace(txt_madt.Text) || string.IsNullOrWhiteSpace(txt_tendt.Text)
                || string.IsNullOrWhiteSpace(cb_khoa.Text) || string.IsNullOrWhiteSpace(txt_linhvuc.Text) || string.IsNullOrWhiteSpace(cbo_tiendo.Text))
                return false;

            return true;
        }
        public bool KtraMaDT(string ma)
        {
            try
            {
                string sql = "select * from DeTai where MaDeTai = '" + ma + "'";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }


            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã đề tài !", "Thông báo");
            }
            return true;
        }

        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Ktra())
            {
                if (KtraMaDT(txt_madt.Text))
                {
                    try
                    {
                        string date1 = dtp_ngaybd.Value.ToString("yyyy/MM/dd");
                        string date2 = dtp_ngaykt.Value.ToString("yyyy/MM/dd");
                        string cap = "Cấp Cơ Sở";
                        string sql = "insert into DeTai values('" + txt_madt.Text + "',N'" + txt_tendt.Text + "',N'" + cb_khoa.Text + "',N'" + txt_linhvuc.Text + "',N'" + cap + "',N'Giảng viên')";
                        int up = my.Update(sql);
                        if (up > 0)
                        {
                            string sql1 = "insert into TienDoDeTai values('" + txt_madt.Text + "','" + date1 + "','" + date2 + "',N'" + cbo_tiendo.Text + "')";
                            int up1 = my.Update(sql1);
                            if (up1 > 0)
                            {
                                MessageBox.Show("Thêm thông tin thành công ", "Thông báo");
                                LoadDL();
                                txt_madt.Clear();
                                txt_tendt.Clear();
                                txt_linhvuc.Clear();
                                cb_khoa.SelectedIndex = -1;
                                cbo_tiendo.SelectedIndex = -1;
                            }
                            else
                            {
                                MessageBox.Show("Thêm thông tin không thành công {2}", "Thông báo");
                            }

                        }
                        else
                        {
                            MessageBox.Show("Thêm thông tin không thành công {1}", "Thông báo");
                        }




                    }
                    catch
                    {
                        MessageBox.Show("Lỗi ! không thêm thành công ", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show("Đã có mã đề tài này .Vui lòng nhập lại !", "Thông báo");
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
                if (!KtraMaDT(txt_madt.Text))
                {
                    try
                    {
                        string cap = "Cấp Cơ Sở";
                        string date1 = dtp_ngaybd.Value.ToString("yyyy/MM/dd");
                        string date2 = dtp_ngaykt.Value.ToString("yyyy/MM/dd");
                        string sql = "update DeTai set TenDeTai = N'" + txt_tendt.Text + "',Khoa = N'" + cb_khoa.Text + "',LinhVuc=N'" + txt_linhvuc.Text + "',Capdetai = N'" + cap + "' ,DoiTuong = N'Giảng viên' where MaDeTai = '" + txt_madt.Text + "' ";
                        int up = my.Update(sql);
                        if (up > 0)
                        {
                            string sql1 = "update TienDoDeTai set NgayBatDau = '" + date1 + "',NgayKetThuc = '" + date2 + "',TienDo=N'" + cbo_tiendo.Text + "' where MaDeTai = '" + txt_madt.Text + "' ";
                            int up1 = my.Update(sql1);
                            if (up1 > 0)
                            {
                                MessageBox.Show("Sửa thông tin thành công ", "Thông báo");
                                LoadDL();
                                txt_madt.Clear();
                                txt_tendt.Clear();
                                txt_linhvuc.Clear();
                                cb_khoa.SelectedIndex = -1;
                                cbo_tiendo.SelectedIndex = -1;
                            }
                            else
                            {
                                MessageBox.Show("Sửa thông tin không thành công {2}", "Thông báo");
                            }


                        }
                        else
                        {
                            MessageBox.Show("Sửa thông tin không thành công {1}", "Thông báo");
                        }





                    }
                    catch
                    {
                        MessageBox.Show("Lỗi ! không Sửa thành công ", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show("Không có mã đề tài này .Vui lòng nhập lại !", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn đề tài muốn sửa !", "Thông báo");
            }
        }

        private void btn_xoa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Ktra())
            {
                if (!KtraMaDT(txt_madt.Text))
                {
                    try
                    {
                        DialogResult tb = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                        if (tb == DialogResult.OK)
                        {

                            string ma = txt_madt.Text;

                            int up2 = my.Update("delete from HoiDong where MaDeTai = '" + ma + "' ");
                            int up3 = my.Update("delete from ChiTietGVDeTai where MaDeTai = '" + ma + "' ");
                            int up4 = my.Update("delete from BanBaoCao where MaDeTai = '" + ma + "' ");
                            int up5 = my.Update("delete from CTDeTai where MaDeTai = '" + ma + "' ");
                            int up6 = my.Update("delete from TienDoDeTai where MaDeTai = '" + ma + "' ");
                            int up7 = my.Update("delete from DeTai where MaDeTai = '" + ma + "' ");

                            if (up2 >= 0)
                            {
                                if (up3 >= 0)
                                {
                                    if (up4 >= 0)
                                    {
                                        if (up5 >= 0)
                                        {
                                            if (up6 >= 0)
                                            {
                                                if (up7 > 0)
                                                {
                                                    MessageBox.Show("Xóa thông tin thành công ", "Thông báo");
                                                    LoadDL();
                                                    LoadDLGV(ma);
                                                    LoadDLHD(ma);

                                                    txt_madt.Clear();
                                                    txt_tendt.Clear();
                                                    txt_linhvuc.Clear();
                                                    cb_khoa.SelectedIndex = -1;
                                                    cbo_tiendo.SelectedIndex = -1;
                                                }
                                                else
                                                {
                                                    MessageBox.Show("Xóa thông tin không thành công {7} ", "Thông báo");

                                                }
                                            }
                                            else
                                            {
                                                MessageBox.Show("Xóa thông tin không thành công {6} ", "Thông báo");

                                            }
                                        }
                                        else
                                        {
                                            MessageBox.Show("Xóa thông tin không thành công {5} ", "Thông báo");

                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Xóa thông tin không thành công {4} ", "Thông báo");

                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Xóa thông tin không thành công {3} ", "Thông báo");

                                }
                            }
                            else
                            {
                                MessageBox.Show("Xóa thông tin không thành công {2} ", "Thông báo");

                            }

                        }
                        else
                        {

                        }


                    }
                    catch
                    {
                        MessageBox.Show("Lỗi ! không Sửa thành công ", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show("Không có mã đề tài này .Vui lòng nhập lại !", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Vui chọn đề tài muốn xóa !", "Thông báo");
            }
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
                    string cap = "Cấp Cơ Sở";
                    if (cbo_tk.Text == "Mã Đề Tài")
                    {
                        try
                        {
                            string sql = @"SELECT DeTai.MaDeTai, DeTai.TenDeTai, DeTai.Khoa, DeTai.LinhVuc, TienDoDeTai.NgayBatDau, TienDoDeTai.NgayKetThuc, TienDoDeTai.TienDo
                                            FROM DeTai
                                            JOIN TienDoDeTai ON TienDoDeTai.MaDeTai = DeTai.MaDeTai
                                            WHERE DeTai.CapDeTai = N'" + cap + "' AND DeTai.MaDeTai LIKE '%" + txt_timkiem.Text + "%' AND DeTai.DoiTuong = N'Giảng viên' ";

                            DataTable dt = my.DocDL(sql);
                            dgv_dt.DataSource = dt;
                            dgv_dt.Columns[0].HeaderText = "Mã đề tài";
                            dgv_dt.Columns[1].HeaderText = "Tên đề tài";
                            dgv_dt.Columns[2].HeaderText = "Khoa";
                            dgv_dt.Columns[3].HeaderText = "Lĩnh vực";
                            dgv_dt.Columns[4].HeaderText = "Ngày bắt đầu";
                            dgv_dt.Columns[5].HeaderText = "Ngày kết thúc";
                            dgv_dt.Columns[6].HeaderText = "Tiến độ";

                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm theo mã đề tài  !", "Thông báo");
                        }
                    }
                    else if (cbo_tk.Text == "Tên Đề Tài")
                    {
                        try
                        {

                            string sql = @"SELECT DeTai.MaDeTai, DeTai.TenDeTai, DeTai.Khoa, DeTai.LinhVuc, TienDoDeTai.NgayBatDau, TienDoDeTai.NgayKetThuc, TienDoDeTai.TienDo
                                            FROM DeTai
                                            JOIN TienDoDeTai ON TienDoDeTai.MaDeTai = DeTai.MaDeTai
                                            WHERE DeTai.CapDeTai = N'" + cap + "' AND DeTai.TenDeTai LIKE N'%" + txt_timkiem.Text + "%' AND DeTai.DoiTuong = N'Giảng viên' ";

                            DataTable dt = my.DocDL(sql);
                            dgv_dt.DataSource = dt;
                            dgv_dt.Columns[0].HeaderText = "Mã đề tài";
                            dgv_dt.Columns[1].HeaderText = "Tên đề tài";
                            dgv_dt.Columns[2].HeaderText = "Khoa";
                            dgv_dt.Columns[3].HeaderText = "Lĩnh vực";
                            dgv_dt.Columns[4].HeaderText = "Ngày bắt đầu";
                            dgv_dt.Columns[5].HeaderText = "Ngày kết thúc";
                            dgv_dt.Columns[6].HeaderText = "Tiến độ";

                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm theo tên đề tài  !", "Thông báo");
                        }
                    }
                    else
                    {
                        try
                        {

                            string sql = @"SELECT DeTai.MaDeTai, DeTai.TenDeTai, DeTai.Khoa, DeTai.LinhVuc, TienDoDeTai.NgayBatDau, TienDoDeTai.NgayKetThuc, TienDoDeTai.TienDo
                                            FROM DeTai
                                            JOIN TienDoDeTai ON TienDoDeTai.MaDeTai = DeTai.MaDeTai
                                            WHERE DeTai.CapDeTai = N'" + cap + "' AND DeTai.Khoa LIKE N'%" + txt_timkiem.Text + "%'  AND DeTai.DoiTuong = N'Giảng viên' ";

                            DataTable dt = my.DocDL(sql);
                            dgv_dt.DataSource = dt;
                            dgv_dt.Columns[0].HeaderText = "Mã đề tài";
                            dgv_dt.Columns[1].HeaderText = "Tên đề tài";
                            dgv_dt.Columns[2].HeaderText = "Khoa";
                            dgv_dt.Columns[3].HeaderText = "Lĩnh vực";
                            dgv_dt.Columns[4].HeaderText = "Ngày bắt đầu";
                            dgv_dt.Columns[5].HeaderText = "Ngày kết thúc";
                            dgv_dt.Columns[6].HeaderText = "Tiến độ";

                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm theo khoa  !", "Thông báo");
                        }
                    }

                }
            }
        }

        private void btn_refresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadDL();
            txt_madt.Clear();
            txt_tendt.Clear();
            txt_linhvuc.Clear();
            txt_timkiem.Clear();
            cb_khoa.SelectedIndex = -1;
            cbo_tk.SelectedIndex = -1;
            txt_tengv.Clear();
            cbo_chucvugv.SelectedIndex = -1;
            txt_tengvhd.Clear();
            cbo_chucvuhd.SelectedIndex = -1;

            dgv_gv.DataSource = null;
            dgv_hd.DataSource = null;

            cbo_tiendo.SelectedIndex = -1;
            cbo_magvhd.SelectedIndex = -1;

            cbo_gv.SelectedIndex = -1;
        }


        public void MaGiangVien()
        {
            try
            {
                cbo_gv.Items.Clear();
                string sql = "select * from GiangVien ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    for (int i = 0; i < tb.Rows.Count; i++)
                    {
                        string ma = tb.Rows[i]["MaGV"].ToString();
                        cbo_gv.Items.Add(ma);
                    }
                }
            }
            catch
            {
                MessageBox.Show("lỗi lấy dữ liệu mã giảng viên !", "Thông báo");
            }
        }
        public void MaGiangVienHD()
        {
            try
            {
                cbo_magvhd.Items.Clear();
                string sql = "select * from GiangVien ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    for (int i = 0; i < tb.Rows.Count; i++)
                    {
                        string ma = tb.Rows[i]["MaGV"].ToString();
                        cbo_magvhd.Items.Add(ma);
                    }
                }
            }
            catch
            {
                MessageBox.Show("lỗi lấy dữ liệu mã giảng viên hội dồng!", "Thông báo");
            }
        }

        private void dgv_dt_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_madt.Text = dgv_dt.CurrentRow.Cells[0].Value.ToString();
                txt_tendt.Text = dgv_dt.CurrentRow.Cells[1].Value.ToString();
                cb_khoa.Text = dgv_dt.CurrentRow.Cells[2].Value.ToString();
                txt_linhvuc.Text = dgv_dt.CurrentRow.Cells[3].Value.ToString();
                dtp_ngaybd.Text = dgv_dt.CurrentRow.Cells[4].Value.ToString();
                dtp_ngaykt.Text = dgv_dt.CurrentRow.Cells[5].Value.ToString();
                cbo_tiendo.Text = dgv_dt.CurrentRow.Cells[6].Value.ToString();

                if (e.RowIndex >= 0)
                {



                    object cellValue = dgv_dt.Rows[e.RowIndex].Cells[0].Value;
                    string madt = cellValue.ToString();
                    try
                    {


                        Madt = madt;
                        LoadDLGV(madt);
                        LoadDLHD(madt);
                        MaGiangVien();
                        MaGiangVienHD();


                    }
                    catch
                    {
                        MessageBox.Show("Lỗi hiển thị thông tin thành viên tham gia đề tài !", "Thông báo");
                    }

                }

            }
            catch
            {
                MessageBox.Show("Lỗi lấy dữ liệu lên textbox !", "Thông báo");
            }
        }

        private void dgv_gv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                cbo_gv.Text = dgv_gv.CurrentRow.Cells[0].Value.ToString();
                txt_tengv.Text = dgv_gv.CurrentRow.Cells[1].Value.ToString();
                cbo_chucvugv.Text = dgv_gv.CurrentRow.Cells[2].Value.ToString();

            }
            catch
            {
                MessageBox.Show("lỗi hiển thị dữ liệu giảng viên!", "Thông báo");
            }
        }

        private void dgv_hd_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                cbo_magvhd.Text = dgv_hd.CurrentRow.Cells[0].Value.ToString();
                txt_tengvhd.Text = dgv_hd.CurrentRow.Cells[1].Value.ToString();
                cbo_chucvuhd.Text = dgv_hd.CurrentRow.Cells[2].Value.ToString();

            }
            catch
            {
                MessageBox.Show("lỗi hiển thị dữ liệu hội đồng!", "Thông báo");
            }
        }

        private void cbo_gv_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cbo_gv.SelectedIndex != -1)
                {
                    string ma = cbo_gv.SelectedItem.ToString();
                    if (!string.IsNullOrWhiteSpace(ma))
                    {
                        string sql = "select HoTen from GiangVien where MaGV = '" + ma + "' ";
                        DataTable tb = my.DocDL(sql);
                        if (tb.Rows.Count > 0)
                        {
                            string hoten = tb.Rows[0][0].ToString();
                            txt_tengv.Text = hoten;
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("lỗi hiển thị tên giảng viên!", "Thông báo");
            }
        }




        public bool KtraMaGVDT(string magv, string madt)
        {
            try
            {
                string sql = "select * from ChiTietGVDeTai where MaGV = '" + magv + "' and MaDeTai = '" + madt + "' ";
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

        public bool KtraMaHDNT(string mahd, string madt)
        {
            try
            {
                string sql = "select * from HoiDongNgoaiTruong where MaHD = '" + mahd + "' and MaDeTai = '" + madt + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }
            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã hội đồng !", "Thông báo");
            }
            return true;
        }



        public bool KtraMaGVHD(string magv, string madt)
        {
            try
            {
                string sql = "select * from HoiDong where MaGV = '" + magv + "' and MaDeTai = '" + madt + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }
            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra mã giảng viên hội đồng!", "Thông báo");
            }
            return true;
        }

        public bool KtraDT(string mahd)
        {
            try
            {
                string sql = "select * from ChiTietGVDeTai where MaGV = '" + mahd + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }
            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra giảng viên tham gia đề tài !", "Thông báo");
            }
            return true;
        }

        public bool KtraHD(string mahd)
        {
            try
            {
                string sql = "select * from HoiDong where MaGV = '" + mahd + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }
            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra giảng viên tham gia hội đồng !", "Thông báo");
            }
            return true;
        }

        private void btn_joingv_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrWhiteSpace(Madt))
            {
                MessageBox.Show("Vui lòng chọn đề tài ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(cbo_gv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text) || string.IsNullOrWhiteSpace(cbo_chucvugv.Text))
                {
                    MessageBox.Show("Vui lòng chọn đầy đủ thông tin !", "Thông báo");
                }
                else
                {
                    try
                    {
                        if (string.IsNullOrWhiteSpace(txt_madt.Text) )
                        {
                            MessageBox.Show("Vui lòng chọn đề tài muốn tham gia !", "Thông báo");
                        }
                        else
                        {
                            string madt = Madt;
                            if (KtraMaGVDT(cbo_gv.Text, madt))
                            {

                                if (KtraHD(cbo_gv.Text))
                                {
                                    string sql = "insert into ChiTietGVDeTai values ('" + cbo_gv.Text + "','" + madt + "',N'" + cbo_chucvugv.Text + "')";
                                    int up = my.Update(sql);
                                    if (up > 0)
                                    {
                                        MessageBox.Show("Tham gia đề tài thành công !", "Thông báo");
                                        LoadDLGV(madt);
                                        txt_tengv.Clear();
                                        cbo_gv.SelectedIndex = -1;
                                        cbo_chucvugv.SelectedIndex = -1;
                                    }
                                    else
                                    {
                                        MessageBox.Show("Tham gia đề tài không thành công !", "Thông báo");
                                    }



                                }
                                else
                                {
                                    MessageBox.Show("Giảng viên đã tham gia hội đồng !", "Thông báo");
                                }
                            }
                            else
                            {
                                MessageBox.Show("Giảng viên đã tham gia đề tài !", "Thông báo");
                            }
                        }

                    }
                    catch
                    {
                        MessageBox.Show("Lỗi tham gia đề tài !", "Thông báo");
                    }
                }
            }
        }

        private void btn_cancelgv_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Madt))
            {
                MessageBox.Show("Vui lòng chọn đề tài ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(cbo_gv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text) || string.IsNullOrWhiteSpace(cbo_chucvugv.Text))
                {
                    MessageBox.Show("Vui lòng chọn đầy đủ thông tin !", "Thông báo");
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(txt_madt.Text))
                    {
                        MessageBox.Show("Vui lòng chọn đề tài muốn loại bỏ thành viên !", "Thông báo");
                    }
                    else
                    {
                        try
                        {
                            string madt = Madt;
                            string sql = "delete from ChiTietGVDeTai where MaGV='" + cbo_gv.Text + "' and MaDeTai = '" + madt + "' ";
                            int up = my.Update(sql);
                            if (up > 0)
                            {
                                MessageBox.Show("Loại bỏ thành viên thành công !", "Thông báo");
                                LoadDLGV(madt);
                                txt_tengv.Clear();
                                cbo_gv.SelectedIndex = -1;
                                cbo_chucvugv.SelectedIndex = -1;
                            }
                            else
                            {
                                MessageBox.Show("Loại bỏ thành viên không thành công !", "Thông báo");
                            }


                        }
                        catch
                        {
                            MessageBox.Show("Lỗi loại bỏ thành viên đề tài !", "Thông báo");
                        }
                    }

                }
            }
        }

        private void btn_suagv_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Madt))
            {
                MessageBox.Show("Vui lòng chọn đề tài ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(cbo_gv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text) || string.IsNullOrWhiteSpace(cbo_chucvugv.Text))
                {
                    MessageBox.Show("Vui lòng chọn đầy đủ thông tin !", "Thông báo");
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(txt_madt.Text))
                    {
                        MessageBox.Show("Vui lòng chọn đề tài muốn sửa thành viên !", "Thông báo");
                    }
                    else
                    {
                        try
                        {
                            string madt = Madt;
                            string sql = "update ChiTietGVDeTai set ChucVu = N'" + cbo_chucvugv.Text + "' where MaGV='" + cbo_gv.Text + "' and MaDeTai = '" + madt + "' ";
                            int up = my.Update(sql);
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa thành viên thành công !", "Thông báo");
                                LoadDLGV(madt);
                                txt_tengv.Clear();
                                cbo_gv.SelectedIndex = -1;
                                cbo_chucvugv.SelectedIndex = -1;

                            }
                            else
                            {
                                MessageBox.Show("Sửa thành viên không thành công !", "Thông báo");
                            }

                        }
                        catch
                        {
                            MessageBox.Show("Lỗi sửa thành viên đề tài !", "Thông báo");
                        }
                    }

                }
            }
        }

        private void cbo_magvhd_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cbo_magvhd.SelectedIndex != -1)
                {
                    string ma = cbo_magvhd.SelectedItem.ToString();
                    if (!string.IsNullOrWhiteSpace(ma))
                    {
                        string sql = "select HoTen from GiangVien where MaGV = '" + ma + "' ";
                        DataTable tb = my.DocDL(sql);
                        if (tb.Rows.Count > 0)
                        {
                            string hoten = tb.Rows[0][0].ToString();
                            txt_tengvhd.Text = hoten;
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("lỗi hiển thị tên giảng viên hội đồng!", "Thông báo");
            }
        }

        private void btn_joinhd_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Madt))
            {
                MessageBox.Show("Vui lòng chọn đề tài ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(cbo_magvhd.Text) || string.IsNullOrWhiteSpace(txt_tengvhd.Text) || string.IsNullOrWhiteSpace(cbo_chucvuhd.Text))
                {
                    MessageBox.Show("Vui lòng chọn đầy đủ thông tin !", "Thông báo");
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(txt_madt.Text))
                    {
                        MessageBox.Show("Vui lòng chọn đề tài muốn tham gia !", "Thông báo");
                    }
                    else
                    {
                        try
                        {
                            string madt = Madt;
                            if (KtraDT(cbo_magvhd.Text))
                            {

                                if (KtraMaGVHD(cbo_magvhd.Text, madt))
                                {
                                    string sql = "insert into HoiDong values ('" + cbo_magvhd.Text + "','" + madt + "',N'" + cbo_chucvuhd.Text + "')";
                                    int up = my.Update(sql);
                                    if (up > 0)
                                    {
                                        MessageBox.Show("Tham gia hội đồng thành công !", "Thông báo");
                                        LoadDLHD(madt);
                                        txt_tengvhd.Clear();
                                        cbo_magvhd.SelectedIndex = -1;
                                        cbo_chucvuhd.SelectedIndex = -1;
                                    }
                                    else
                                    {
                                        MessageBox.Show("Tham gia hội đồng không thành công !", "Thông báo");
                                    }


                                }
                                else
                                {
                                    MessageBox.Show("Giảng viên đã tham gia hội đồng !", "Thông báo");
                                }
                            }
                            else
                            {
                                MessageBox.Show("Giảng viên này đã tham gia đề tài !", "Thông báo");
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tham gia hội đồng !", "Thông báo");
                        }
                    }

                }
            }
        }

        private void btn_cancelhd_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Madt))
            {
                MessageBox.Show("Vui lòng chọn đề tài ", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(txt_madt.Text))
                {
                    MessageBox.Show("Vui lòng chọn đề tài muốn loại bỏ thành viên !", "Thông báo");
                }
                else
                {
                    try
                    {
                        string madt = Madt;
                        string sql = "delete from HoiDong where MaGV='" + cbo_magvhd.Text + "' and MaDeTai = '" + madt + "' ";
                        int up = my.Update(sql);
                        if (up > 0)
                        {
                            MessageBox.Show("Loại bỏ thành viên hội đồng thành công !", "Thông báo");
                            LoadDLHD(madt);
                            txt_tengvhd.Clear();
                            cbo_magvhd.SelectedIndex = -1;
                            cbo_chucvuhd.SelectedIndex = -1;
                        }
                        else
                        {
                            MessageBox.Show("Loại bỏ thành viên hội đồng không thành công !", "Thông báo");
                        }


                    }
                    catch
                    {
                        MessageBox.Show("Lỗi loại bỏ thành viên hội đồng !", "Thông báo");
                    }
                }

            }
        
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

                head.Value2 = "DANH SÁCH ĐỀ TÀI NGHIÊN CỨU KHOA HỌC CẤP CƠ SỞ";

                head.Font.Bold = true;

                head.Font.Name = "Times New Roman";

                head.Font.Size = "20";

                head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range cl1 = oSheet.get_Range("A3", "A3");
                cl1.Value = "Mã đề tài";

                Excel.Range cl2 = oSheet.get_Range("B3", "B3");
                cl2.Value = "Tên đề tài";
                Excel.Range cl3 = oSheet.get_Range("C3", "C3");
                cl3.Value = "Khoa";

                Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                cl4.Value = "Lĩnh vực";

                Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                cl5.Value = "Ngày bắt đầu";

                Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                cl6.Value = "Ngày kết thúc";

                Excel.Range cl7 = oSheet.get_Range("G3", "G3");
                cl7.Value = "Tiến độ";



                Excel.Range rowHead = oSheet.get_Range("A3", "G3");
                rowHead.Font.Bold = true;
                rowHead.Font.Size = 13;
                rowHead.Font.Name = "Times New Roman";
                rowHead.Borders.LineStyle = Excel.Constants.xlSolid;
                rowHead.Interior.ColorIndex = 6;
                rowHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                int line = 4;
                for (int i = 0; i < dgv_dt.Rows.Count - 1; i++)
                {
                    Excel.Range line1 = oSheet.get_Range("A" + (line + i).ToString(), "A" + (line + i).ToString());
                    line1.Value = dgv_dt.Rows[i].Cells[0].Value.ToString();
                    line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line1.Borders.LineStyle = Excel.Constants.xlSolid;
                    line1.Font.Name = "Times New Roman";

                    Excel.Range line2 = oSheet.get_Range("B" + (line + i).ToString(), "B" + (line + i).ToString());
                    line2.Value = dgv_dt.Rows[i].Cells[1].Value.ToString();
                    line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line2.Borders.LineStyle = Excel.Constants.xlSolid;
                    line2.Font.Name = "Times New Roman";

                    Excel.Range line3 = oSheet.get_Range("C" + (line + i).ToString(), "C" + (line + i).ToString());
                    line3.Value = dgv_dt.Rows[i].Cells[2].Value.ToString();
                    line3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line3.Borders.LineStyle = Excel.Constants.xlSolid;
                    line3.Font.Name = "Times New Roman";

                    Excel.Range line4 = oSheet.get_Range("D" + (line + i).ToString(), "D" + (line + i).ToString());
                    line4.Value = dgv_dt.Rows[i].Cells[3].Value.ToString();
                    line4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line4.Borders.LineStyle = Excel.Constants.xlSolid;
                    line4.Font.Name = "Times New Roman";

                    Excel.Range line5 = oSheet.get_Range("E" + (line + i).ToString(), "E" + (line + i).ToString());
                    line5.Value = dgv_dt.Rows[i].Cells[4].Value.ToString();
                    line5.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line5.Borders.LineStyle = Excel.Constants.xlSolid;
                    line5.Font.Name = "Times New Roman";

                    Excel.Range line6 = oSheet.get_Range("F" + (line + i).ToString(), "F" + (line + i).ToString());
                    line6.Value = dgv_dt.Rows[i].Cells[5].Value.ToString();
                    line6.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line6.Borders.LineStyle = Excel.Constants.xlSolid;
                    line6.Font.Name = "Times New Roman";

                    Excel.Range line7 = oSheet.get_Range("G" + (line + i).ToString(), "G" + (line + i).ToString());
                    line7.Value = dgv_dt.Rows[i].Cells[6].Value.ToString();
                    line7.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line7.Borders.LineStyle = Excel.Constants.xlSolid;
                    line7.Font.Name = "Times New Roman";

                }


                oSheet.Name = "DTCCS";
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

        private void barButtonItem4_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ExcelExport();
        }

        public void excelCT()
        {
            try
            {

                DataTable dataTable = LayDuLieuBaoCao();


                Excel.Application oExcel = new Excel.Application();
                Excel.Workbook workbook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet oSheet = (Excel.Worksheet)workbook.Worksheets[1];




                Excel.Range head = oSheet.get_Range("A1", "I1");

                head.MergeCells = true;

                head.Value2 = "DANH SÁCH CHI TIẾT ĐỀ TÀI NGHIÊN CỨU KHOA HỌC CẤP CƠ SỞ  ";

                head.Font.Bold = true;

                head.Font.Name = "Times New Roman";

                head.Font.Size = "20";

                head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range cl1 = oSheet.get_Range("A3", "A3");
                cl1.Value = "Mã đề tài";

                Excel.Range cl2 = oSheet.get_Range("B3", "B3");
                cl2.Value = "Tên đề tài";

                Excel.Range cl3 = oSheet.get_Range("C3", "C3");
                cl3.Value = "Khoa";

                Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                cl4.Value = "Lĩnh vực";

                Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                cl5.Value = "Ngày bắt đầu";

                Excel.Range cl10 = oSheet.get_Range("F3", "F3");
                cl10.Value = "Ngày kết thúc";

                Excel.Range cl6 = oSheet.get_Range("G3", "G3");
                cl6.Value = "Tiến độ";

                Excel.Range cl7 = oSheet.get_Range("H3", "H3");
                cl7.Value = "Thành viên nghiên cứu - Chức vụ";

                Excel.Range cl8 = oSheet.get_Range("I3", "I3");
                cl8.Value = "Thành viên hội đồng - Chức vụ";





                Excel.Range rowHead = oSheet.get_Range("A3", "I3");
                rowHead.Font.Bold = true;
                rowHead.Font.Name = "Times New Roman";
                rowHead.Font.Size = 13;
                rowHead.Borders.LineStyle = Excel.Constants.xlSolid;
                rowHead.Interior.ColorIndex = 6;
                rowHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Sau đó, thêm dữ liệu từ DataTable
                int line = 4;
                int lines = 4;
                string ma;
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {

                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        oSheet.Cells[i + line, j + 1] = dataTable.Rows[i][j];
                        oSheet.Cells[i + line, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        oSheet.Cells[i + line, j + 1].Borders.LineStyle = Excel.Constants.xlSolid;
                        oSheet.Cells[i + line, j + 1].Font.Name = "Times New Roman";

                    }


                    ma = dataTable.Rows[i][0].ToString();
                    //
                    string query = "SELECT GiangVien.HoTen, ChiTietGVDeTai.ChucVu FROM ChiTietGVDeTai, GiangVien WHERE ChiTietGVDeTai.MaGV = GiangVien.MaGV AND MaDeTai = '" + ma + "' ";

                    DataTable dt = my.DocDL(query);

                    Excel.Range line1 = oSheet.get_Range("H" + (lines).ToString(), "H" + (lines).ToString());

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {

                        string cel = dt.Rows[row]["HoTen"].ToString() + "-" + dt.Rows[row]["ChucVu"].ToString() + "\n";
                        line1.Value += cel;


                    }
                    line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line1.Borders.LineStyle = Excel.Constants.xlSolid;
                    line1.Font.Name = "Times New Roman";
                    //


                    //
                    string hd = "SELECT GiangVien.HoTen, HoiDong.ChucVuHD FROM HoiDong,GiangVien WHERE HoiDong.MaGV = GiangVien.MaGV AND MaDeTai = '" + ma + "' ";

                    DataTable dthd = my.DocDL(hd);

                    Excel.Range line2 = oSheet.get_Range("I" + (lines).ToString(), "I" + (lines).ToString());

                    for (int row = 0; row < dthd.Rows.Count; row++)
                    {

                        string cel = dthd.Rows[row]["HoTen"].ToString() + "-" + dthd.Rows[row]["ChucVuHD"].ToString() + "\n";
                        line2.Value += cel;


                    }
                    line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line2.Borders.LineStyle = Excel.Constants.xlSolid;
                    line2.Font.Name = "Times New Roman";
                    //



                    lines++;




                }

                oSheet.Name = "CTDTCCS";
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


        public DataTable LayDuLieuBaoCao()
        {

            string sql = @"  
            SELECT 
                DeTai.MaDeTai,
                DeTai.TenDeTai,
                DeTai.Khoa,
                DeTai.LinhVuc,
	            TienDoDeTai.NgayBatDau,
	            TienDoDeTai.NgayKetThuc,
	            TienDoDeTai.TienDo               
            FROM 
                DeTai
            LEFT JOIN
                TienDoDeTai ON DeTai.MaDeTai = TienDoDeTai.MaDeTai            
            WHERE DeTai.Capdetai = N'Cấp Cơ Sở' and DeTai.DoiTuong = N'Giảng viên'
           ";
            DataTable dataTable = my.DocDL(sql);

            return dataTable;
        }

        private void StartLongTask(frm_please formWaiting)
        {
            Thread.Sleep(3000);
            formWaiting.Invoke(new Action(() => formWaiting.Close()));
        }

        private void barButtonItem5_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            frm_please formWaiting = new frm_please();
            formWaiting.StartPosition = FormStartPosition.CenterScreen;

            Thread thread = new Thread(() => StartLongTask(formWaiting));
            thread.Start();

            formWaiting.ShowDialog();
            excelCT();
        }
        public void excelCT1DT()
        {
            try
            {
                if (Ktra())
                {
                    DataTable dataTable = LayDuLieuBaoCao1DT();


                    Excel.Application oExcel = new Excel.Application();
                    Excel.Workbook workbook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    Excel.Worksheet oSheet = (Excel.Worksheet)workbook.Worksheets[1];




                    Excel.Range head = oSheet.get_Range("A1", "I1");

                    head.MergeCells = true;

                    head.Value2 = "THÔNG TIN ĐỀ TÀI NGHIÊN CỨU KHOA HỌC CẤP CƠ SỞ  ";

                    head.Font.Bold = true;

                    head.Font.Name = "Times New Roman";

                    head.Font.Size = "20";

                    head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    Excel.Range cl1 = oSheet.get_Range("A3", "A3");
                    cl1.Value = "Mã đề tài";

                    Excel.Range cl2 = oSheet.get_Range("B3", "B3");
                    cl2.Value = "Tên đề tài";

                    Excel.Range cl3 = oSheet.get_Range("C3", "C3");
                    cl3.Value = "Khoa";

                    Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                    cl4.Value = "Lĩnh vực";

                    Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                    cl5.Value = "Ngày bắt đầu";

                    Excel.Range cl10 = oSheet.get_Range("F3", "F3");
                    cl10.Value = "Ngày kết thúc";

                    Excel.Range cl6 = oSheet.get_Range("G3", "G3");
                    cl6.Value = "Tiến độ";

                    Excel.Range cl7 = oSheet.get_Range("H3", "H3");
                    cl7.Value = "Thành viên nghiên cứu - Chức vụ";

                    Excel.Range cl8 = oSheet.get_Range("I3", "I3");
                    cl8.Value = "Thành viên hội đồng - Chức vụ";





                    Excel.Range rowHead = oSheet.get_Range("A3", "I3");
                    rowHead.Font.Bold = true;
                    rowHead.Font.Name = "Times New Roman";
                    rowHead.Font.Size = 13;
                    rowHead.Borders.LineStyle = Excel.Constants.xlSolid;
                    rowHead.Interior.ColorIndex = 6;
                    rowHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    // Sau đó, thêm dữ liệu từ DataTable
                    int line = 4;
                    int lines = 4;
                    string ma;
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {

                        for (int j = 0; j < dataTable.Columns.Count; j++)
                        {
                            oSheet.Cells[i + line, j + 1] = dataTable.Rows[i][j];
                            oSheet.Cells[i + line, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            oSheet.Cells[i + line, j + 1].Borders.LineStyle = Excel.Constants.xlSolid;
                            oSheet.Cells[i + line, j + 1].Font.Name = "Times New Roman";

                        }


                        ma = dataTable.Rows[i][0].ToString();
                        //
                        string query = "SELECT GiangVien.HoTen, ChiTietGVDeTai.ChucVu FROM ChiTietGVDeTai, GiangVien WHERE ChiTietGVDeTai.MaGV = GiangVien.MaGV AND MaDeTai = '" + ma + "' ";

                        DataTable dt = my.DocDL(query);

                        Excel.Range line1 = oSheet.get_Range("H" + (lines).ToString(), "H" + (lines).ToString());

                        for (int row = 0; row < dt.Rows.Count; row++)
                        {

                            string cel = dt.Rows[row]["HoTen"].ToString() + "-" + dt.Rows[row]["ChucVu"].ToString() + "\n";
                            line1.Value += cel;


                        }
                        line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line1.Borders.LineStyle = Excel.Constants.xlSolid;
                        line1.Font.Name = "Times New Roman";
                        //


                        //
                        string hd = "SELECT GiangVien.HoTen, HoiDong.ChucVuHD FROM HoiDong,GiangVien WHERE HoiDong.MaGV = GiangVien.MaGV AND MaDeTai = '" + ma + "' ";

                        DataTable dthd = my.DocDL(hd);

                        Excel.Range line2 = oSheet.get_Range("I" + (lines).ToString(), "I" + (lines).ToString());

                        for (int row = 0; row < dthd.Rows.Count; row++)
                        {

                            string cel = dthd.Rows[row]["HoTen"].ToString() + "-" + dthd.Rows[row]["ChucVuHD"].ToString() + "\n";
                            line2.Value += cel;


                        }
                        line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line2.Borders.LineStyle = Excel.Constants.xlSolid;
                        line2.Font.Name = "Times New Roman";
                        //



                        lines++;




                    }

                    oSheet.Name = "DTCCS";
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
                    MessageBox.Show($"Vui lòng chọn đề tài muốn export dữ liệu", "Thông báo");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi xuất báo cáo: {ex.Message}");
            }
        }
        public DataTable LayDuLieuBaoCao1DT()
        {

            string sql = @"  
            SELECT 
                DeTai.MaDeTai,
                DeTai.TenDeTai,
                DeTai.Khoa,
                DeTai.LinhVuc,
	            TienDoDeTai.NgayBatDau,
	            TienDoDeTai.NgayKetThuc,
	            TienDoDeTai.TienDo               
            FROM 
                DeTai
            LEFT JOIN
                TienDoDeTai ON DeTai.MaDeTai = TienDoDeTai.MaDeTai            
            WHERE DeTai.Capdetai = N'Cấp Cơ Sở' and DeTai.DoiTuong = N'Giảng viên' and DeTai.MaDeTai = '" + txt_madt.Text + "' ";
            DataTable dataTable = my.DocDL(sql);

            return dataTable;
        }
        private void barButtonItem6_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            excelCT1DT();
        }

        private void btn_suahd_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrWhiteSpace(Madt))
            {
                MessageBox.Show("Vui lòng chọn đề tài", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(cbo_magvhd.Text) || string.IsNullOrWhiteSpace(txt_tengvhd.Text) || string.IsNullOrWhiteSpace(cbo_chucvuhd.Text))
                {
                    MessageBox.Show("Vui lòng chọn đầy đủ thông tin !", "Thông báo");
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(txt_madt.Text))
                    {
                        MessageBox.Show("Vui lòng đề tài muốn sửa thành viên !", "Thông báo");
                    }
                    else
                    {
                        try
                        {
                            string madt = Madt;
                            string sql = "update HoiDong set ChucVuHD = N'" + cbo_chucvuhd.Text + "' where MaGV='" + cbo_magvhd.Text + "' and MaDeTai = '" + madt + "' ";
                            int up = my.Update(sql);
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa thành viên thành công !", "Thông báo");
                                LoadDLHD(madt);
                                txt_tengvhd.Clear();
                                cbo_magvhd.SelectedIndex = -1;
                                cbo_chucvuhd.SelectedIndex = -1;
                            }
                            else
                            {
                                MessageBox.Show("Sửa thành viên không thành công !", "Thông báo");

                            }


                        }
                        catch
                        {
                            MessageBox.Show("Lỗi sửa thành viên đề tài !", "Thông báo");
                        }
                    }

                }
            }
        }

        private void dgv_dt_CellBorderStyleChanged(object sender, EventArgs e)
        {

        }


        private void LoadProductList()
        {
            try
            {

                productList = new List<string>();
                string query = "SELECT MaGV FROM GiangVien";
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
                MessageBox.Show($"Lỗi thực hiện tạo danh sách giảng viên", "Lỗi");
            }

        }
        private void ShowSuggestions(List<string> suggestions)
        {
            list_gv.Items.Clear();
            list_gv.Items.AddRange(suggestions.ToArray());

            list_gv.Visible = suggestions.Any();
        }

        private void ShowSuggestionsHD(List<string> suggestions)
        {
            list_hd.Items.Clear();
            list_hd.Items.AddRange(suggestions.ToArray());

            list_hd.Visible = suggestions.Any();
        }
        private void cbo_gv_TextChanged(object sender, EventArgs e)
        {
            if (dgv_gv.Rows.Count >= 0)
            {
                string searchTerm = cbo_gv.Text.ToLower();
                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
                    List<string> filteredProducts = productList
                   .Where(product => product.ToLower().Contains(searchTerm))
                   .ToList();

                    ShowSuggestions(filteredProducts);
                }
                else
                {
                    list_gv.Visible = false;
                }


            }
            else
            {
                MessageBox.Show($"Vui lòng chọn đề tài", "Thông báo");
            }
        }

        private void list_gv_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_gv.SelectedItem != null)
            {
                string selectedProduct = list_gv.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    cbo_gv.Text = selectedProduct;
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

        private void list_hd_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_hd.SelectedItem != null)
            {
                string selectedProduct = list_hd.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    cbo_magvhd.Text = selectedProduct;
                    list_hd.Visible = false;
                    string sql = "select HoTen from GiangVien where MaGV = '" + selectedProduct + "' ";
                    DataTable tb = my.DocDL(sql);
                    if (tb.Rows.Count > 0)
                    {
                        string hoten = tb.Rows[0][0].ToString();
                        txt_tengvhd.Text = hoten;
                    }
                }

            }
        }

        private void cbo_gv_TextChanged_1(object sender, EventArgs e)
        {
            if (dgv_gv.Rows.Count >= 0)
            {
                string searchTerm = cbo_gv.Text.ToLower();
                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
                    List<string> filteredProducts = productList
                   .Where(product => product.ToLower().Contains(searchTerm))
                   .ToList();

                    ShowSuggestions(filteredProducts);
                }
                else
                {
                    list_gv.Visible = false;
                }


            }
            else
            {
                MessageBox.Show($"Vui lòng chọn đề tài", "Thông báo");
            }
        }

        private void cbo_magvhd_TextChanged(object sender, EventArgs e)
        {
            if (dgv_hd.Rows.Count >= 0)
            {
                string searchTerm = cbo_magvhd.Text.ToLower();
                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
                    List<string> filteredProducts = productList
                   .Where(product => product.ToLower().Contains(searchTerm))
                   .ToList();

                    ShowSuggestionsHD(filteredProducts);
                }
                else
                {
                    list_hd.Visible = false;
                }


            }
            else
            {
                MessageBox.Show($"Vui lòng chọn đề tài", "Thông báo");
            }
        }
    }
}
