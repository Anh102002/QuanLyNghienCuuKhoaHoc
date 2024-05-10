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
    public partial class uc_capnhatcuocthi : DevExpress.XtraEditors.XtraUserControl
    {
        MyClass my = new MyClass();
        private List<string> productListCT;
        private List<string> productListDT;
        private List<string> productListGV;
        private List<string> productListKM;

        private string mact;
        private string madt;
        public uc_capnhatcuocthi()
        {
            InitializeComponent();
        }
        public string Mact
        {
            get { return this.mact; }
            set { this.mact = value; }
        }

        public string Madt
        {
            get { return this.madt; }
            set { this.madt = value; }
        }
        public void loadDL()
        {
            try
            {

               
                string query = @" select KetQuaCuocThi.MaCuocThi,CuocThiSTKN.TenCuocThi,KetQuaCuocThi.NgayThanhLap,KetQuaCuocThi.SoQuyetDinh,KetQuaCuocThi.HinhThucToChuc,CuocThiSTKN.CapCT 
                                from KetQuaCuocThi 
                               LEFT JOIN CuocThiSTKN ON CuocThiSTKN.MaCuocThi = KetQuaCuocThi.MaCuocThi
                                ORDER BY CuocThiSTKN.MaCuocThi";
                DataTable dt = my.DocDL(query);
                dgv_ct.DataSource = dt;
                dgv_ct.Columns[0].HeaderText = "Mã cuộc thi";
                dgv_ct.Columns[1].HeaderText = "Tên cuộc thi";
                dgv_ct.Columns[1].Width = 300;
                dgv_ct.Columns[2].HeaderText = "Ngày thành lập";
                dgv_ct.Columns[3].HeaderText = "Số quyết định";
                dgv_ct.Columns[4].HeaderText = "Hình thức tổ chức";
                dgv_ct.Columns[5].HeaderText = "Cấp cuộc thi";





            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu cuộc thi ", "Lỗi");
            }
        }

        public void loadDLDT(string ma)
        {
            try
            {

                string query = " select MaDoi,TenDoi,TenYTuong,DonVi,GiaiThuong from DoiThamGiaCuocThi where MaCuocThi = '" + ma + "' ";
                DataTable dt = my.DocDL(query);
                dgv_doithi.DataSource = dt;
                dgv_doithi.Columns[0].HeaderText = "Mã đội thi";
                dgv_doithi.Columns[1].HeaderText = "Tên đội thi";
                dgv_doithi.Columns[2].HeaderText = "Tên ý tưởng";
                dgv_doithi.Columns[3].HeaderText = "Đơn vị";
                dgv_doithi.Columns[4].HeaderText = "Giải thưởng";
                dgv_doithi.Columns[1].Width = 200;

            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu đội thi ", "Lỗi");
            }
        }

        public void loadDLBGK(string ma)
        {
            try
            {

                string query = @" select BGKCuocThi.MaGV,GiangVien.HoTen,BGKCuocThi.ChucVu,BGKCuocThi.VaiTro 
                                    from BGKCuocThi,GiangVien where BGKCuocThi.MaCuocThi = '" + ma+ "'  and BGKCuocThi.MaGV = GiangVien.MaGV ";
                DataTable dt = my.DocDL(query);
                dgv_bgk.DataSource = dt;
                dgv_bgk.Columns[0].HeaderText = "Mã giảng viên";
                dgv_bgk.Columns[1].HeaderText = "Tên giảng viên";
                dgv_bgk.Columns[2].HeaderText = "Chức vụ";
                dgv_bgk.Columns[3].HeaderText = "Vai trò";

            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu thành viên ban giám khảo ", "Lỗi");
            }
        }
        public void loadDLBGKNT(string ma)
        {
            try
            {

                string query = @"select BGKCuocThiKM.MaKM,TVNgoaiTruong.HoTen,BGKCuocThiKM.ChucVu,BGKCuocThiKM.VaiTro  
                                    from BGKCuocThiKM
                                    LEFT JOIN TVNgoaiTruong on  BGKCuocThiKM.MaKM = TVNgoaiTruong.MaKM
                                    where BGKCuocThiKM.MaCuocThi = '" + ma + "' ";
                DataTable dt = my.DocDL(query);
                dgv_bgknt.DataSource = dt;
                dgv_bgknt.Columns[0].HeaderText = "Mã thành viên";
                dgv_bgknt.Columns[1].HeaderText = "Tên thành viên";
                dgv_bgknt.Columns[2].HeaderText = "Chức vụ";
                dgv_bgknt.Columns[3].HeaderText = "Vai trò";

            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị dữ liệu thành viên ban giám khảo ngoài trường ", "Lỗi");
            }
        }
        private void btn_export_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

        }

        private void uc_capnhatcuocthi_Load(object sender, EventArgs e)
        {
            try
            {
                loadDL();
                LoadProductListDT();
                LoadProductListCT();
                LoadProductListGV();
                LoadProductListKM();
                
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
                dtp_ngaytl.Text = dgv_ct.CurrentRow.Cells[2].Value.ToString();
                txt_soquetdinh.Text = dgv_ct.CurrentRow.Cells[3].Value.ToString();
                cbo_hinhthuc.Text = dgv_ct.CurrentRow.Cells[4].Value.ToString();
                


                
                
                    Mact = txt_mact.Text;
                    string ma = Mact;
                    loadDLDT(ma);
                    loadDLBGK(ma);
                    loadDLBGKNT(ma);
                
                
                

            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị thông tin cuộc thi", "Lỗi");
            }
        }

        private void dgv_doithi_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_madoithi.Text = dgv_doithi.CurrentRow.Cells[0].Value.ToString();
                txt_tendoithi.Text = dgv_doithi.CurrentRow.Cells[1].Value.ToString();
                txt_ytuong.Text = dgv_doithi.CurrentRow.Cells[2].Value.ToString();
                cbo_giaithuong.Text = dgv_doithi.CurrentRow.Cells[4].Value.ToString();

                Madt = txt_madoithi.Text;
            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị thông tin đội thi", "Lỗi");
            }
        }

        private void dgv_bgk_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_magv.Text = dgv_bgk.CurrentRow.Cells[0].Value.ToString();
                txt_tengv.Text = dgv_bgk.CurrentRow.Cells[1].Value.ToString();
                txt_chucvugv.Text = dgv_bgk.CurrentRow.Cells[2].Value.ToString();
                cbo_vaitroBGK.Text = dgv_bgk.CurrentRow.Cells[3].Value.ToString();
                

            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị thông tin ban giám khảo", "Lỗi");
            }
        }

        private void dgv_bgknt_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_matv.Text = dgv_bgknt.CurrentRow.Cells[0].Value.ToString();
                txt_tentv.Text = dgv_bgknt.CurrentRow.Cells[1].Value.ToString();
                txt_chucvuBGK.Text = dgv_bgknt.CurrentRow.Cells[2].Value.ToString();
                cbo_vaitroBGKNT.Text = dgv_bgknt.CurrentRow.Cells[3].Value.ToString();
                string ma = txt_madoithi.Text;

            }
            catch
            {
                MessageBox.Show("$ Lỗi hiển thị thông tin ban giám khảo ngoài trường", "Lỗi");
            }
        }


        private void LoadProductListCT()
        {
            try
            {

                productListCT = new List<string>();
                string query = "SELECT MaCuocThi FROM CuocThiSTKN";
                DataTable tb = my.DocDL(query);
                if (tb.Rows.Count > 0)
                {
                    for (int i = 0; i < tb.Rows.Count; i++)
                    {
                        string ma = tb.Rows[i][0].ToString();
                        productListCT.Add(ma);
                    }
                }
            }
            catch
            {
                MessageBox.Show($"Lỗi thực hiện tạo danh sách mã cuộc thi", "Lỗi");
            }

        }
        private void ShowSuggestionsCT(List<string> suggestions)
        {
            list_ct.Items.Clear();
            list_ct.Items.AddRange(suggestions.ToArray());

            list_ct.Visible = suggestions.Any();
        }

        private void txt_mact_TextChanged(object sender, EventArgs e)
        {
            string searchTerm = txt_mact.Text.ToLower();
            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                List<string> filteredProducts = productListCT
               .Where(product => product.ToLower().Contains(searchTerm))
               .ToList();

                if (filteredProducts != null)
                {
                    ShowSuggestionsCT(filteredProducts);
                }
                else
                {
                    list_ct.Visible = false;

                }


            }
            else
            {
                list_ct.Visible = false;
                txt_tenct.Clear();
            }
        }

        private void list_ct_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_ct.SelectedItem != null)
            {
                string selectedProduct = list_ct.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_mact.Text = selectedProduct;
                    list_ct.Visible = false;
                    string sql = "select TenCuocThi from CuocThiSTKN where MaCuocThi = '" + selectedProduct + "' ";
                    DataTable tb = my.DocDL(sql);
                    if (tb.Rows.Count > 0)
                    {
                        string hoten = tb.Rows[0][0].ToString();
                        txt_tenct.Text = hoten;
                    }

                }

            }
        }
        private void LoadProductListDT()
        {
            try
            {

                productListDT = new List<string>();
                string query = "SELECT MaDoi FROM DoiThamGiaCuocThi";
                DataTable tb = my.DocDL(query);
                if (tb.Rows.Count > 0)
                {
                    for (int i = 0; i < tb.Rows.Count; i++)
                    {
                        string ma = tb.Rows[i][0].ToString();
                        productListDT.Add(ma);
                    }
                }
            }
            catch
            {
                MessageBox.Show($"Lỗi thực hiện tạo danh sách mã đội thi", "Lỗi");
            }

        }
        //private void ShowSuggestionsDT(List<string> suggestions)
        //{
        //    list_dt.Items.Clear();
        //    list_dt.Items.AddRange(suggestions.ToArray());

        //    list_dt.Visible = suggestions.Any();
        //}
        private void txt_madoithi_TextChanged(object sender, EventArgs e)
        {
            //string searchTerm = txt_madoithi.Text.ToLower();
            //if (!string.IsNullOrWhiteSpace(searchTerm))
            //{
            //    List<string> filteredProducts = productListDT
            //   .Where(product => product.ToLower().Contains(searchTerm))
            //   .ToList();

            //    if (filteredProducts != null)
            //    {
            //        ShowSuggestionsDT(filteredProducts);
            //    }
            //    else
            //    {
            //        list_dt.Visible = false;

            //    }


            //}
            //else
            //{
            //    list_dt.Visible = false;
            //    txt_tendoithi.Clear();
            //}
        }

        private void list_dt_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (list_dt.SelectedItem != null)
            //{
            //    string selectedProduct = list_dt.SelectedItem.ToString();
            //    if (!string.IsNullOrWhiteSpace(selectedProduct))
            //    {
            //        txt_madoithi.Text = selectedProduct;
            //        list_dt.Visible = false;
            //        string sql = "select TenDoi,TenYTuong from DoiThamGiaCuocThi where MaDoi = '" + selectedProduct + "' ";
            //        DataTable tb = my.DocDL(sql);
            //        if (tb.Rows.Count > 0)
            //        {
            //            string hoten = tb.Rows[0][0].ToString();
            //            string ytuong = tb.Rows[0][1].ToString();
            //            txt_tendoithi.Text = hoten;
            //            txt_ytuong.Text = ytuong;
            //        }

            //    }

            //}
        }
        //
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
                MessageBox.Show($"Lỗi thực hiện tạo danh sách mã ban giám khảo", "Lỗi");
            }

        }
        private void ShowSuggestionsGV(List<string> suggestions)
        {
            listBgk.Items.Clear();
            listBgk.Items.AddRange(suggestions.ToArray());

            listBgk.Visible = suggestions.Any();
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
                    listBgk.Visible = false;

                }


            }
            else
            {
                listBgk.Visible = false;
                txt_tengv.Clear();
            }
        }

        private void listBgk_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBgk.SelectedItem != null)
            {
                string selectedProduct = listBgk.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_magv.Text = selectedProduct;
                    listBgk.Visible = false;
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
        ///////////
        ///

        private void LoadProductListKM()
        {
            try
            {

                productListKM = new List<string>();
                string query = "SELECT MaKM FROM TVNgoaiTruong ";
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
                MessageBox.Show($"Lỗi thực hiện tạo danh sách mã ban giám khảo ngoài trường", "Lỗi");
            }

        }

        private void ShowSuggestionsKM(List<string> suggestions)
        {
            list_bgknt.Items.Clear();
            list_bgknt.Items.AddRange(suggestions.ToArray());

            list_bgknt.Visible = suggestions.Any();
        }

        private void txt_matv_TextChanged(object sender, EventArgs e)
        {
            string searchTerm = txt_matv.Text.ToLower();
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
                    list_bgknt.Visible = false;

                }


            }
            else
            {
                list_bgknt.Visible = false;
                txt_tentv.Clear();
            }
        }

        private void list_bgknt_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_bgknt.SelectedItem != null)
            {
                string selectedProduct = list_bgknt.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_matv.Text = selectedProduct;
                    list_bgknt.Visible = false;
                    string sql = "select HoTen from TVNgoaiTruong where MaKM = '" + selectedProduct + "' ";
                    DataTable tb = my.DocDL(sql);
                    if (tb.Rows.Count > 0)
                    {
                        string hoten = tb.Rows[0][0].ToString();

                        txt_tentv.Text = hoten;

                    }

                }

            }
        }

        public bool KiemTraNULL()
        {
            if (string.IsNullOrWhiteSpace(txt_mact.Text) || string.IsNullOrWhiteSpace(txt_soquetdinh.Text)
                || string.IsNullOrWhiteSpace(txt_tenct.Text) || string.IsNullOrWhiteSpace(cbo_hinhthuc.Text) || string.IsNullOrWhiteSpace(dtp_ngaytl.Text))
            {
                return false;
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

        public bool KtraMaKQCT(string ma)
        {
            try
            {
                string sql = "select * from KetQuaCuocThi where MaCuocThi = '" + ma + "'";
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
        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if(KiemTraNULL())
            {
                if(!KtraMaCT(txt_mact.Text))
                {
                    if(KtraMaKQCT(txt_mact.Text))
                    {
                        try
                        {
                            string ngaytl = dtp_ngaytl.Value.ToString("yyyy/MM/dd");

                            string sql = "insert into KetQuaCuocThi values (@Ma,@Ngay,@SoQD,@HinhThuc) ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Ma", txt_mact.Text);
                            command.Parameters.AddWithValue("@Ngay", ngaytl);
                            command.Parameters.AddWithValue("@SoQD", txt_soquetdinh.Text);
                            command.Parameters.AddWithValue("@HinhThuc", cbo_hinhthuc.Text);


                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Thêm thông tin thành công", "Thông báo");

                                txt_tenct.Clear();
                                txt_mact.Clear();
                                txt_soquetdinh.Clear();

                                cbo_hinhthuc.SelectedIndex = -1;

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
                        MessageBox.Show("Đã có cuộc thi này rồi ", "Thông báo");
                    }
                    
                }
            }
            else
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
            }
        }

        private void btn_sua_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (KiemTraNULL())
            {
                if (!KtraMaCT(txt_mact.Text))
                {
                    if (!KtraMaKQCT(txt_mact.Text))
                    {
                        try
                        {
                            string ngaytl = dtp_ngaytl.Value.ToString("yyyy/MM/dd");
                            string sql = "update KetQuaCuocThi set NgayThanhLap=@Ngay,SoQuyetDinh=@SoQD,HinhThucToChuc=@HinhThuc where MaCuocThi=@Ma ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Ma", txt_mact.Text);
                            command.Parameters.AddWithValue("@Ngay", ngaytl);
                            command.Parameters.AddWithValue("@SoQD", txt_soquetdinh.Text);
                            command.Parameters.AddWithValue("@HinhThuc", cbo_hinhthuc.Text);


                            int up = command.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa thông tin thành công", "Thông báo");

                                txt_tenct.Clear();
                                txt_mact.Clear();
                                txt_soquetdinh.Clear();

                                cbo_hinhthuc.SelectedIndex = -1;

                                loadDL();

                            }



                        }
                        catch
                        {
                            MessageBox.Show("Lỗi ! không sửa thành công ", "Lỗi");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Không có cuộc thi này ", "Thông báo");
                    }

                }
            }
            else
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
            }
        }

        private void btn_xoa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (KiemTraNULL())
            {
                if (!KtraMaCT(txt_mact.Text))
                {
                    if (!KtraMaKQCT(txt_mact.Text))
                    {
                        try
                        {
                            DialogResult result = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                            if (result == DialogResult.OK)
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
                                        int dem = 0;
                                        string query = "select MaDoi from DoiThamGiaCuocThi where  MaCuocThi = '" + txt_mact.Text + "'";
                                        DataTable tb = my.DocDL(query);
                                        if (tb.Rows.Count > 0)
                                        {
                                            for (int i = 0; i < tb.Rows.Count; i++)
                                            {
                                                string madoi = tb.Rows[i][0].ToString();

                                                string btt = "delete from BaiThuyetTrinhCT where MaDoi = @MaDoi ";

                                                SqlCommand commandBTT = my.SqlCommand(btt);
                                                commandBTT.Parameters.AddWithValue("@MaDoi", madoi);

                                                int upBTT = commandBTT.ExecuteNonQuery();
                                                if (upBTT >= 0)
                                                {
                                                    string suaGT = "update DoiThamGiaCuocThi set GiaiThuong=@giai where MaDoi = @MaDoi ";
                                                    SqlCommand commandGT = my.SqlCommand(suaGT);
                                                    commandGT.Parameters.AddWithValue("@giai", "");
                                                    commandGT.Parameters.AddWithValue("@MaDoi", madoi);
                                                    int sua = commandGT.ExecuteNonQuery();
                                                    if (sua >= 0)
                                                    {
                                                        dem++;
                                                    }
                                                }
                                            }
                                        }
                                        if (dem > 0)
                                        {

                                            string sql = "delete from KetQuaCuocThi where MaCuocThi=@Ma3 ";
                                            SqlCommand command = my.SqlCommand(sql);
                                            command.Parameters.AddWithValue("@Ma3", txt_mact.Text);



                                            int up = command.ExecuteNonQuery();
                                            if (up > 0)
                                            {
                                                MessageBox.Show("Xóa thông tin thành công", "Thông báo");

                                                txt_tenct.Clear();
                                                txt_mact.Clear();
                                                txt_soquetdinh.Clear();
                                                cbo_hinhthuc.SelectedIndex = -1;
                                                //
                                                txt_tendoithi.Clear();
                                                txt_ytuong.Clear();
                                                txt_madoithi.Clear();
                                                cbo_giaithuong.SelectedIndex = -1;
                                                //
                                                txt_tengv.Clear();
                                                txt_magv.Clear();
                                                txt_chucvugv.Clear();
                                                cbo_vaitroBGK.SelectedIndex = -1;
                                                //
                                                txt_tentv.Clear();
                                                txt_matv.Clear();
                                                txt_chucvuBGK.Clear();
                                                cbo_vaitroBGKNT.SelectedIndex = -1;
                                                //
                                                loadDL();
                                                dgv_doithi.DataSource = null;
                                                dgv_bgk.DataSource = null;
                                                dgv_bgknt.DataSource = null;

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
                            MessageBox.Show("Lỗi ! không xóa thành công{" + ex.Message + "} ", "Lỗi");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Không có cuộc thi này ", "Thông báo");
                    }
                    
                }
                
            }
            else
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
            }
        }

        private void btn_refesh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            cbo_tk.SelectedIndex = -1;
            txt_timkiem.Clear();

            txt_tenct.Clear();
            txt_mact.Clear();            
            txt_soquetdinh.Clear();
            cbo_hinhthuc.SelectedIndex = -1;

            txt_tendoithi.Clear();
            txt_ytuong.Clear();
            txt_madoithi.Clear();
            cbo_giaithuong.SelectedIndex = -1;

            loadDL();

            txt_tengv.Clear();
            txt_magv.Clear();
            cbo_vaitroBGK.SelectedIndex = -1;
            txt_chucvugv.Clear();

            
            txt_tentv.Clear();
            txt_matv.Clear();
            txt_chucvuBGK.Clear();
            cbo_vaitroBGKNT.SelectedIndex = -1;

            dgv_doithi.DataSource = null;
            dgv_bgk.DataSource = null;
            dgv_bgknt.DataSource = null;
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
                            string query = @" select KetQuaCuocThi.MaCuocThi,CuocThiSTKN.TenCuocThi,KetQuaCuocThi.NgayThanhLap,KetQuaCuocThi.SoQuyetDinh,KetQuaCuocThi.HinhThucToChuc,CuocThiSTKN.CapCT 
                                from KetQuaCuocThi 
                               LEFT JOIN CuocThiSTKN ON CuocThiSTKN.MaCuocThi = KetQuaCuocThi.MaCuocThi where KetQuaCuocThi.MaCuocThi like '%"+txt_timkiem.Text+"%' ";
                            DataTable dt = my.DocDL(query);
                            dgv_ct.DataSource = dt;
                            dgv_ct.Columns[0].HeaderText = "Mã cuộc thi";
                            dgv_ct.Columns[1].HeaderText = "Tên cuộc thi";
                            dgv_ct.Columns[1].Width = 300;
                            dgv_ct.Columns[2].HeaderText = "Ngày thành lập";
                            dgv_ct.Columns[3].HeaderText = "Số quyết định";
                            dgv_ct.Columns[4].HeaderText = "Hình thức tổ chức";
                            dgv_ct.Columns[5].HeaderText = "Cấp cuộc thi";


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

                            string query = @" select KetQuaCuocThi.MaCuocThi,CuocThiSTKN.TenCuocThi,KetQuaCuocThi.NgayThanhLap,KetQuaCuocThi.SoQuyetDinh,KetQuaCuocThi.HinhThucToChuc,CuocThiSTKN.CapCT 
                                from KetQuaCuocThi 
                               LEFT JOIN CuocThiSTKN ON CuocThiSTKN.MaCuocThi = KetQuaCuocThi.MaCuocThi where CuocThiSTKN.TenCuocThi like N'%" + txt_timkiem.Text + "%' ";
                            DataTable dt = my.DocDL(query);
                            dgv_ct.DataSource = dt;
                            dgv_ct.Columns[0].HeaderText = "Mã cuộc thi";
                            dgv_ct.Columns[1].HeaderText = "Tên cuộc thi";
                            dgv_ct.Columns[1].Width = 300;
                            dgv_ct.Columns[2].HeaderText = "Ngày thành lập";
                            dgv_ct.Columns[3].HeaderText = "Số quyết định";
                            dgv_ct.Columns[4].HeaderText = "Hình thức tổ chức";
                            dgv_ct.Columns[5].HeaderText = "Cấp cuộc thi";


                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm theo tên cuộc thi !", "Thông báo");
                        }
                    }


                }
            }
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                string query = @" select KetQuaCuocThi.MaCuocThi,CuocThiSTKN.TenCuocThi,KetQuaCuocThi.NgayThanhLap,KetQuaCuocThi.SoQuyetDinh,KetQuaCuocThi.HinhThucToChuc,CuocThiSTKN.CapCT 
                                from KetQuaCuocThi 
                               LEFT JOIN CuocThiSTKN ON CuocThiSTKN.MaCuocThi = KetQuaCuocThi.MaCuocThi where CuocThiSTKN.CapCT = N'Cấp Quốc Gia' or CuocThiSTKN.CapCT = N'Cấp Bộ' ";
                DataTable dt = my.DocDL(query);
                dgv_ct.DataSource = dt;
                dgv_ct.Columns[0].HeaderText = "Mã cuộc thi";
                dgv_ct.Columns[1].HeaderText = "Tên cuộc thi";
                dgv_ct.Columns[1].Width = 300;
                dgv_ct.Columns[2].HeaderText = "Ngày thành lập";
                dgv_ct.Columns[3].HeaderText = "Số quyết định";
                dgv_ct.Columns[4].HeaderText = "Hình thức tổ chức";
                dgv_ct.Columns[5].HeaderText = "Cấp cuộc thi";


            }
            catch
            {
                MessageBox.Show("Lỗi hiển thị cuộc thị các quốc gia, cấp bộ  !", "Thông báo");
            }
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                string query = @" select KetQuaCuocThi.MaCuocThi,CuocThiSTKN.TenCuocThi,KetQuaCuocThi.NgayThanhLap,KetQuaCuocThi.SoQuyetDinh,KetQuaCuocThi.HinhThucToChuc,CuocThiSTKN.CapCT 
                                from KetQuaCuocThi 
                               LEFT JOIN CuocThiSTKN ON CuocThiSTKN.MaCuocThi = KetQuaCuocThi.MaCuocThi where CuocThiSTKN.CapCT = N'Cấp Trường' ";
                DataTable dt = my.DocDL(query);
                dgv_ct.DataSource = dt;
                dgv_ct.Columns[0].HeaderText = "Mã cuộc thi";
                dgv_ct.Columns[1].HeaderText = "Tên cuộc thi";
                dgv_ct.Columns[1].Width = 300;
                dgv_ct.Columns[2].HeaderText = "Ngày thành lập";
                dgv_ct.Columns[3].HeaderText = "Số quyết định";
                dgv_ct.Columns[4].HeaderText = "Hình thức tổ chức";
                dgv_ct.Columns[5].HeaderText = "Cấp cuộc thi";


            }
            catch
            {
                MessageBox.Show("Lỗi hiển thị cuộc thị các trường  !", "Thông báo");
            }
        }

        private void barButtonItem3_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            loadDL();
        }
        public bool KtraDT()
        {
            if(string.IsNullOrWhiteSpace(txt_madoithi.Text) || string.IsNullOrWhiteSpace(txt_tendoithi.Text) 
                || string.IsNullOrWhiteSpace(txt_ytuong.Text) || string.IsNullOrWhiteSpace(cbo_giaithuong.Text))
            {
                return false;
            }
            return true;
        }
        private void btn_suadoithi_Click(object sender, EventArgs e)
        {
            if(KtraDT())
            {
                try
                {
                    string sql = "update DoiThamGiaCuocThi set GiaiThuong = @GiaiThuong where MaDoi = @Madoi";
                    SqlCommand command = my.SqlCommand(sql);
                    command.Parameters.AddWithValue("@Madoi", txt_madoithi.Text);
                    command.Parameters.AddWithValue("@GiaiThuong",cbo_giaithuong.Text);
                    int up = command.ExecuteNonQuery();
                    if(up > 0)
                    {
                        MessageBox.Show("Cập nhật kết quả đội thi thành công", "Thông báo");
                        txt_madoithi.Clear();
                        txt_tendoithi.Clear();
                        txt_ytuong.Clear();
                        cbo_giaithuong.SelectedIndex = -1;
                        loadDLDT(txt_mact.Text);
                    }
                }
                catch
                {
                    MessageBox.Show("Lỗi cập nhật đội thi", "Lỗi");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin","Thông báo");
            }
        }

        private void btn_baithuyettrinh_Click(object sender, EventArgs e)
        {
            string madoi = Madt;
            frm_baithuyettrinh frm = new frm_baithuyettrinh();
            frm.Madt = madoi;
            frm.Show();
        }

        public DataTable LayDuLieuBaoCao()
        {

            string query = @" select KetQuaCuocThi.MaCuocThi,CuocThiSTKN.TenCuocThi,KetQuaCuocThi.NgayThanhLap,KetQuaCuocThi.SoQuyetDinh,KetQuaCuocThi.HinhThucToChuc,CuocThiSTKN.CapCT 
                                from KetQuaCuocThi 
                               LEFT JOIN CuocThiSTKN ON CuocThiSTKN.MaCuocThi = KetQuaCuocThi.MaCuocThi
                                ORDER BY CuocThiSTKN.MaCuocThi";
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




                Excel.Range head = oSheet.get_Range("A1", "I1");

                head.MergeCells = true;

                head.Value2 = "CHI TIẾT KẾT QUẢ CUỘC THI";

                head.Font.Bold = true;

                head.Font.Name = "Times New Roman";

                head.Font.Size = "20";

                head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range cl1 = oSheet.get_Range("A3", "A3");
                cl1.Value = "Mã cuộc thi";

                Excel.Range cl2 = oSheet.get_Range("B3", "B3");
                cl2.Value = "Tên cuộc thi";

                Excel.Range cl3 = oSheet.get_Range("C3", "C3");
                cl3.Value = "Ngày thành lập";

                Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                cl4.Value = "Số quyết định";

                Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                cl5.Value = "Hình thức tổ chức";

                Excel.Range cl10 = oSheet.get_Range("F3", "F3");
                cl10.Value = "Cấp cuộc thi";

                Excel.Range cl6 = oSheet.get_Range("G3", "G3");
                cl6.Value = "Đội thi";

                Excel.Range cl7 = oSheet.get_Range("H3", "H3");
                cl7.Value = "Ban giám khảo trong trường";

                Excel.Range cl8 = oSheet.get_Range("I3", "I3");
                cl8.Value = "Ban giám khảo ngoài trường";

                //Excel.Range cl9 = oSheet.get_Range("J3", "J3");
                //cl9.Value = "Giảng viên hướng dẫn";





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
                    string query = " select MaDoi,TenDoi,TenYTuong,DonVi,GiaiThuong from DoiThamGiaCuocThi where MaCuocThi = '" + maCT + "' ";

                    DataTable dt = my.DocDL(query);

                    Excel.Range line1 = oSheet.get_Range("G" + (lines).ToString(), "G" + (lines).ToString());
                    Excel.Range line2 = oSheet.get_Range("H" + (lines).ToString(), "H" + (lines).ToString());
                    Excel.Range line3 = oSheet.get_Range("I" + (lines).ToString(), "I" + (lines).ToString());
                    //Excel.Range line4 = oSheet.get_Range("I" + (lines).ToString(), "I" + (lines).ToString());

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        string maDT = dt.Rows[row][0].ToString();

                        string cel = dt.Rows[row]["MaDoi"].ToString() + "-" + dt.Rows[row]["TenDoi"].ToString() + "-" + dt.Rows[row]["TenYTuong"].ToString() + "-" + dt.Rows[row]["DonVi"].ToString() +"-" + dt.Rows[row]["GiaiThuong"].ToString() + "\n";
                        line1.Value += cel;
                        //
                       
                    }
                    line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line1.Borders.LineStyle = Excel.Constants.xlSolid;
                    line1.Font.Name = "Times New Roman";
                    //

                    //
                    string bgk = @" select BGKCuocThi.MaGV,GiangVien.HoTen,BGKCuocThi.ChucVu,BGKCuocThi.VaiTro 
                                    from BGKCuocThi,GiangVien where BGKCuocThi.MaCuocThi = '" + maCT + "'  and BGKCuocThi.MaGV = GiangVien.MaGV ";

                    DataTable dt_bgk = my.DocDL(bgk);



                    for (int r = 0; r < dt_bgk.Rows.Count; r++)
                    {

                        string celSV = dt_bgk.Rows[r]["MaGV"].ToString() + "-" + dt_bgk.Rows[r]["HoTen"].ToString() + "-" + dt_bgk.Rows[r]["ChucVu"].ToString() +"-"+ dt_bgk.Rows[r]["VaiTro"].ToString() + "\n";
                        line2.Value += celSV;


                    }
                    line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line2.Borders.LineStyle = Excel.Constants.xlSolid;
                    line2.Font.Name = "Times New Roman";

                    //
                    //
                    string bgknt = @"select BGKCuocThiKM.MaKM,TVNgoaiTruong.HoTen,BGKCuocThiKM.ChucVu,BGKCuocThiKM.VaiTro  
                                    from BGKCuocThiKM
                                    LEFT JOIN TVNgoaiTruong on  BGKCuocThiKM.MaKM = TVNgoaiTruong.MaKM
                                    where BGKCuocThiKM.MaCuocThi = '" + maCT + "' ";

                    DataTable dt_bgknt = my.DocDL(bgknt);



                    for (int r1 = 0; r1 < dt_bgknt.Rows.Count; r1++)
                    {

                        string celSVNT = dt_bgknt.Rows[r1]["MaKM"].ToString() + "-" + dt_bgknt.Rows[r1]["HoTen"].ToString() + "-" + dt_bgknt.Rows[r1]["ChucVu"].ToString() +"-"+ dt_bgknt.Rows[r1]["VaiTro"].ToString() + "\n";
                        line3.Value += celSVNT;


                    }
                    line3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line3.Borders.LineStyle = Excel.Constants.xlSolid;

                    line3.Font.Name = "Times New Roman";

                    //
                    //
                    //string hdnt = @"SELECT GVHDCuocThi.MaDoi, GiangVien.HoTen 
                    //                    FROM GVHDCuocThi, GiangVien 
                    //                    WHERE GVHDCuocThi.MaGV = GiangVien.MaGV AND GVHDCuocThi.MaDoi = '" + maDT + "' ";
                    //DataTable gvhd = my.DocDL(hdnt);



                    //for (int gv = 0; gv < gvhd.Rows.Count; gv++)
                    //{

                    //    string celGV = gvhd.Rows[gv]["MaDoi"].ToString() + "-" + gvhd.Rows[gv]["HoTen"].ToString() + "\n";
                    //    line4.Value += celGV;


                    //}
                    //line4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    //line4.Borders.LineStyle = Excel.Constants.xlSolid;
                    //line4.Font.Name = "Times New Roman";

                    //
                    //
                    lines++;




                }

                oSheet.Name = "CTKQCT";
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
        public DataTable LayDuLieuBaoCao1CT()
        {

            string query = @" select KetQuaCuocThi.MaCuocThi,CuocThiSTKN.TenCuocThi,KetQuaCuocThi.NgayThanhLap,KetQuaCuocThi.SoQuyetDinh,KetQuaCuocThi.HinhThucToChuc,CuocThiSTKN.CapCT 
                                from KetQuaCuocThi 
                               LEFT JOIN CuocThiSTKN ON CuocThiSTKN.MaCuocThi = KetQuaCuocThi.MaCuocThi
                                
                                where KetQuaCuocThi.MaCuocThi = '"+txt_mact.Text+"' ";
            DataTable dataTable = my.DocDL(query);

            return dataTable;
        }
        public void excelCT1CT()
        {
            try
            {
                if(string.IsNullOrWhiteSpace(txt_mact.Text))
                {
                    MessageBox.Show("Vui lòng chọn cuộc thi cần exprort dữ liệu");
                }
                else
                {
                    DataTable dataTable = LayDuLieuBaoCao1CT();


                    Excel.Application oExcel = new Excel.Application();
                    Excel.Workbook workbook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    Excel.Worksheet oSheet = (Excel.Worksheet)workbook.Worksheets[1];




                    Excel.Range head = oSheet.get_Range("A1", "I1");

                    head.MergeCells = true;

                    head.Value2 = "CHI TIẾT KẾT QUẢ CUỘC THI";

                    head.Font.Bold = true;

                    head.Font.Name = "Times New Roman";

                    head.Font.Size = "20";

                    head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    Excel.Range cl1 = oSheet.get_Range("A3", "A3");
                    cl1.Value = "Mã cuộc thi";

                    Excel.Range cl2 = oSheet.get_Range("B3", "B3");
                    cl2.Value = "Tên cuộc thi";

                    Excel.Range cl3 = oSheet.get_Range("C3", "C3");
                    cl3.Value = "Ngày thành lập";

                    Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                    cl4.Value = "Số quyết định";

                    Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                    cl5.Value = "Hình thức tổ chức";

                    Excel.Range cl10 = oSheet.get_Range("F3", "F3");
                    cl10.Value = "Cấp cuộc thi";

                    Excel.Range cl6 = oSheet.get_Range("G3", "G3");
                    cl6.Value = "Đội thi";

                    Excel.Range cl7 = oSheet.get_Range("H3", "H3");
                    cl7.Value = "Ban giám khảo trong trường";

                    Excel.Range cl8 = oSheet.get_Range("I3", "I3");
                    cl8.Value = "Ban giám khảo ngoài trường";

                    //Excel.Range cl9 = oSheet.get_Range("J3", "J3");
                    //cl9.Value = "Giảng viên hướng dẫn";





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
                        string query = " select MaDoi,TenDoi,TenYTuong,DonVi,GiaiThuong from DoiThamGiaCuocThi where MaCuocThi = '" + maCT + "' ";

                        DataTable dt = my.DocDL(query);

                        Excel.Range line1 = oSheet.get_Range("G" + (lines).ToString(), "G" + (lines).ToString());
                        Excel.Range line2 = oSheet.get_Range("H" + (lines).ToString(), "H" + (lines).ToString());
                        Excel.Range line3 = oSheet.get_Range("I" + (lines).ToString(), "I" + (lines).ToString());
                        //Excel.Range line4 = oSheet.get_Range("I" + (lines).ToString(), "I" + (lines).ToString());

                        for (int row = 0; row < dt.Rows.Count; row++)
                        {
                            string maDT = dt.Rows[row][0].ToString();

                            string cel = dt.Rows[row]["MaDoi"].ToString() + "-" + dt.Rows[row]["TenDoi"].ToString() + "-" + dt.Rows[row]["TenYTuong"].ToString() + "-" + dt.Rows[row]["DonVi"].ToString() + "-" + dt.Rows[row]["GiaiThuong"].ToString() + "\n";
                            line1.Value += cel;
                            //

                        }
                        line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line1.Borders.LineStyle = Excel.Constants.xlSolid;
                        line1.Font.Name = "Times New Roman";
                        //

                        //
                        string bgk = @" select BGKCuocThi.MaGV,GiangVien.HoTen,BGKCuocThi.ChucVu,BGKCuocThi.VaiTro 
                                    from BGKCuocThi,GiangVien where BGKCuocThi.MaCuocThi = '" + maCT + "'  and BGKCuocThi.MaGV = GiangVien.MaGV ";

                        DataTable dt_bgk = my.DocDL(bgk);



                        for (int r = 0; r < dt_bgk.Rows.Count; r++)
                        {

                            string celSV = dt_bgk.Rows[r]["MaGV"].ToString() + "-" + dt_bgk.Rows[r]["HoTen"].ToString() + "-" + dt_bgk.Rows[r]["ChucVu"].ToString() + "-" + dt_bgk.Rows[r]["VaiTro"].ToString() + "\n";
                            line2.Value += celSV;


                        }
                        line2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line2.Borders.LineStyle = Excel.Constants.xlSolid;
                        line2.Font.Name = "Times New Roman";

                        //
                        //
                        string bgknt = @"select BGKCuocThiKM.MaKM,TVNgoaiTruong.HoTen,BGKCuocThiKM.ChucVu,BGKCuocThiKM.VaiTro  
                                    from BGKCuocThiKM
                                    LEFT JOIN TVNgoaiTruong on  BGKCuocThiKM.MaKM = TVNgoaiTruong.MaKM
                                    where BGKCuocThiKM.MaCuocThi = '" + maCT + "' ";

                        DataTable dt_bgknt = my.DocDL(bgknt);



                        for (int r1 = 0; r1 < dt_bgknt.Rows.Count; r1++)
                        {

                            string celSVNT = dt_bgknt.Rows[r1]["MaKM"].ToString() + "-" + dt_bgknt.Rows[r1]["HoTen"].ToString() + "-" + dt_bgknt.Rows[r1]["ChucVu"].ToString() + "-" + dt_bgknt.Rows[r1]["VaiTro"].ToString() + "\n";
                            line3.Value += celSVNT;


                        }
                        line3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line3.Borders.LineStyle = Excel.Constants.xlSolid;

                        line3.Font.Name = "Times New Roman";

                        //
                        //
                        //string hdnt = @"SELECT GVHDCuocThi.MaDoi, GiangVien.HoTen 
                        //                    FROM GVHDCuocThi, GiangVien 
                        //                    WHERE GVHDCuocThi.MaGV = GiangVien.MaGV AND GVHDCuocThi.MaDoi = '" + maDT + "' ";
                        //DataTable gvhd = my.DocDL(hdnt);



                        //for (int gv = 0; gv < gvhd.Rows.Count; gv++)
                        //{

                        //    string celGV = gvhd.Rows[gv]["MaDoi"].ToString() + "-" + gvhd.Rows[gv]["HoTen"].ToString() + "\n";
                        //    line4.Value += celGV;


                        //}
                        //line4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        //line4.Borders.LineStyle = Excel.Constants.xlSolid;
                        //line4.Font.Name = "Times New Roman";

                        //
                        //
                        lines++;




                    }

                    oSheet.Name = "CTKQCT";
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
        private void barButtonItem5_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            excelCT1CT();
        }

        private void barButtonItem4_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            frm_please formWaiting = new frm_please();
            formWaiting.StartPosition = FormStartPosition.CenterScreen;

            Thread thread = new Thread(() => StartLongTask(formWaiting));
            thread.Start();

            formWaiting.ShowDialog();
            excelCT();
        }
        public bool KtraBGK(string ma)
        {
            string sql = "select * from BGKCuocThi where MaGV = '"+txt_magv.Text+ "' and MaCuocThi = '"+ma+"' ";
            DataTable tb = my.DocDL(sql);
            if(tb.Rows.Count >0)
            {
                return false;
            }
            return true;
        }
        private void btn_joinbgk_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_magv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text)
                || string.IsNullOrWhiteSpace(txt_chucvugv.Text) || string.IsNullOrWhiteSpace(cbo_vaitroBGK.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {
                    if (!KiemTraNULL())
                    {
                        MessageBox.Show("Vui lòng chọn cuộc thi !", "Thông báo");
                    }
                    else
                    {
                        
                        if (KtraBGK(txt_mact.Text))
                        {
                            string sql = "insert into BGKCuocThi values (@Magv,@Mact,@ChucVu,@Vaitro)";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Magv", txt_magv.Text);
                            comd.Parameters.AddWithValue("@Mact", txt_mact.Text);
                            comd.Parameters.AddWithValue("@ChucVu", txt_chucvugv.Text);
                            comd.Parameters.AddWithValue("@Vaitro", cbo_vaitroBGK.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Thêm ban giám khảo thành công !", "Thông báo");

                                txt_tengv.Clear();
                                txt_magv.Clear();
                                txt_chucvugv.Clear();
                                cbo_vaitroBGK.SelectedIndex = -1;
                                loadDLBGK(txt_mact.Text);

                            }
                            





                        }
                        else
                        {
                            MessageBox.Show("Đã có trong ban giám khảo !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi thêm ban giám khảo!", "Lỗi");
                }
            }
        }

        private void btn_cancelbgk_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_magv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text)
                || string.IsNullOrWhiteSpace(txt_chucvugv.Text) || string.IsNullOrWhiteSpace(cbo_vaitroBGK.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {
                    if (!KiemTraNULL())
                    {
                        MessageBox.Show("Vui lòng chọn cuộc thi !", "Thông báo");
                    }
                    else
                    {

                        if (!KtraBGK(txt_mact.Text))
                        {
                            string sql = "delete from BGKCuocThi where MaGV = @Magv and MaCuocThi = @Mact ";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Magv", txt_magv.Text);
                            comd.Parameters.AddWithValue("@Mact", txt_mact.Text);
                            


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa ban giám khảo thành công !", "Thông báo");

                                txt_tengv.Clear();
                                txt_magv.Clear();
                                txt_chucvugv.Clear();
                                cbo_vaitroBGK.SelectedIndex = -1;
                                loadDLBGK(txt_mact.Text);

                            }
                            





                        }
                        else
                        {
                            MessageBox.Show("không có trong ban giám khảo !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi xóa ban giám khảo!", "Lỗi");
                }
            }
        }

        private void btn_suabgk_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_magv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text)
                || string.IsNullOrWhiteSpace(txt_chucvugv.Text) || string.IsNullOrWhiteSpace(cbo_vaitroBGK.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {
                    if (!KiemTraNULL())
                    {
                        MessageBox.Show("Vui lòng chọn cuộc thi !", "Thông báo");
                    }
                    else
                    {

                        if (!KtraBGK(txt_mact.Text))
                        {
                            string sql = "update BGKCuocThi set ChucVu=@ChucVu,VaiTro=@Vaitro where MaGV = @Magv and MaCuocThi = @Mact";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Magv", txt_magv.Text);
                            comd.Parameters.AddWithValue("@Mact", txt_mact.Text);
                            comd.Parameters.AddWithValue("@ChucVu", txt_chucvugv.Text);
                            comd.Parameters.AddWithValue("@Vaitro", cbo_vaitroBGK.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa ban giám khảo thành công !", "Thông báo");

                                txt_tengv.Clear();
                                txt_magv.Clear();
                                txt_chucvugv.Clear();
                                cbo_vaitroBGK.SelectedIndex = -1;
                                loadDLBGK(txt_mact.Text);

                            }






                        }
                        else
                        {
                            MessageBox.Show("không có trong ban giám khảo !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi sửa ban giám khảo!", "Lỗi");
                }
            }
        }
        public bool KtraBGKNT(string ma)
        {
            string sql = "select * from BGKCuocThiKM where MaKM = '" + txt_matv.Text + "' and MaCuocThi = '" + ma + "' ";
            DataTable tb = my.DocDL(sql);
            if (tb.Rows.Count > 0)
            {
                return false;
            }
            return true;
        }
        private void btn_joinbgknt_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_matv.Text) || string.IsNullOrWhiteSpace(txt_tentv.Text)
                || string.IsNullOrWhiteSpace(txt_chucvuBGK.Text) || string.IsNullOrWhiteSpace(cbo_vaitroBGKNT.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {
                    if (!KiemTraNULL())
                    {
                        MessageBox.Show("Vui lòng chọn cuộc thi !", "Thông báo");
                    }
                    else
                    {

                        if (KtraBGKNT(txt_mact.Text))
                        {
                            string sql = "insert into BGKCuocThiKM values (@Makm,@Mact,@ChucVu,@Vaitro)";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Makm", txt_matv.Text);
                            comd.Parameters.AddWithValue("@Mact", txt_mact.Text);
                            comd.Parameters.AddWithValue("@ChucVu", txt_chucvuBGK.Text);
                            comd.Parameters.AddWithValue("@Vaitro", cbo_vaitroBGKNT.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Thêm ban giám khảo thành công !", "Thông báo");

                                txt_tentv.Clear();
                                txt_matv.Clear();
                                txt_chucvuBGK.Clear();
                                cbo_vaitroBGKNT.SelectedIndex = -1;
                                loadDLBGKNT(txt_mact.Text);

                            }






                        }
                        else
                        {
                            MessageBox.Show("Đã có trong ban giám khảo !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi thêm ban giám khảo!", "Lỗi");
                }
            }
        }

        private void btn_cancelbgknt_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_matv.Text) || string.IsNullOrWhiteSpace(txt_tentv.Text)
                || string.IsNullOrWhiteSpace(txt_chucvuBGK.Text) || string.IsNullOrWhiteSpace(cbo_vaitroBGKNT.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {
                    if (!KiemTraNULL())
                    {
                        MessageBox.Show("Vui lòng chọn cuộc thi !", "Thông báo");
                    }
                    else
                    {

                        if (!KtraBGKNT(txt_mact.Text))
                        {
                            string sql = "delete from BGKCuocThiKM where MaKM = @Makm and MaCuocThi=@Mact ";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Makm", txt_matv.Text);
                            comd.Parameters.AddWithValue("@Mact", txt_mact.Text);
                            


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa ban giám khảo thành công !", "Thông báo");

                                txt_tentv.Clear();
                                txt_matv.Clear();
                                txt_chucvuBGK.Clear();
                                cbo_vaitroBGKNT.SelectedIndex = -1;
                                loadDLBGKNT(txt_mact.Text);

                            }






                        }
                        else
                        {
                            MessageBox.Show("không có trong ban giám khảo !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi xóa ban giám khảo!", "Lỗi");
                }
            }
        }

        private void btn_suabgknt_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_matv.Text) || string.IsNullOrWhiteSpace(txt_tentv.Text)
                || string.IsNullOrWhiteSpace(txt_chucvuBGK.Text) || string.IsNullOrWhiteSpace(cbo_vaitroBGKNT.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
            }
            else
            {
                try
                {
                    if (!KiemTraNULL())
                    {
                        MessageBox.Show("Vui lòng chọn cuộc thi !", "Thông báo");
                    }
                    else
                    {

                        if (!KtraBGKNT(txt_mact.Text))
                        {
                            string sql = "update BGKCuocThiKM set ChucVu=@ChucVu,VaiTro=@Vaitro where MaKM = @Makm and MaCuocThi = @Mact";
                            SqlCommand comd = my.SqlCommand(sql);
                            comd.Parameters.AddWithValue("@Makm", txt_matv.Text);
                            comd.Parameters.AddWithValue("@Mact", txt_mact.Text);
                            comd.Parameters.AddWithValue("@ChucVu", txt_chucvuBGK.Text);
                            comd.Parameters.AddWithValue("@Vaitro", cbo_vaitroBGKNT.Text);


                            int up = comd.ExecuteNonQuery();
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa ban giám khảo thành công !", "Thông báo");

                                txt_tentv.Clear();
                                txt_matv.Clear();
                                txt_chucvuBGK.Clear();
                                cbo_vaitroBGKNT.SelectedIndex = -1;
                                loadDLBGKNT(txt_mact.Text);

                            }






                        }
                        else
                        {
                            MessageBox.Show("không có trong ban giám khảo !", "Thông báo");
                        }
                    }

                }
                catch
                {
                    MessageBox.Show("Lỗi sửa ban giám khảo!", "Lỗi");
                }
            }
        }
    }
}
