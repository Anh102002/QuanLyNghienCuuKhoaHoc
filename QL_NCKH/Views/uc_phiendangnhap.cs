using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QL_NCKH
{
    public partial class uc_phiendangnhap : DevExpress.XtraEditors.XtraUserControl
    {
        MyClass my = new MyClass();
        public uc_phiendangnhap()
        {
            InitializeComponent();
        }
        public void Header()
        {
            dgv_dangnhap.Columns[0].HeaderText = "Id";
            dgv_dangnhap.Columns[1].HeaderText = "Username";
            dgv_dangnhap.Columns[2].HeaderText = "Thời gian đăng nhập";
            dgv_dangnhap.Columns[2].Width = 200;
            dgv_dangnhap.Columns[3].HeaderText = "Thời gian đăng xuất";
            dgv_dangnhap.Columns[3].Width = 200;
        }

        public void loadDL()
        {
            try
            {
                string sql = "select * from PhienDangNhap ";
                DataTable tb = my.DocDL(sql);
                
                    dgv_dangnhap.DataSource = tb;
                    Header();
                
                    
            }
            catch
            {
                MessageBox.Show("Lỗi lấy dữ liệu","Thông báo");
            }
        }
        private void uc_phiendangnhap_Load(object sender, EventArgs e)
        {
            loadDL();
        }

        private void dgv_dangnhap_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
            
        }

        private void btn_sua_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if(dgv_dangnhap.SelectedCells.Count > 0)
            {
                if(!string.IsNullOrWhiteSpace(txt_id.Text) || !string.IsNullOrWhiteSpace(txt_user.Text))
                {
                    try
                    {                       
                        
                        string sql = "update PhienDangNhap set Username='',ThoiGianDangNhap='"+dtp_login.Text+"',ThoiGianDangXuat= '" + dtp_logout.Text + "' where Id = '" + txt_id.Text+ "'   ";
                       my.Update(sql);
                       
                            MessageBox.Show("Thông tin được sửa thành công", "Thông báo", MessageBoxButtons.OKCancel);
                       
                        
                    }
                    catch
                    {
                        MessageBox.Show("Lỗi sửa thông tin dữ liệu", "Thông báo", MessageBoxButtons.OKCancel);

                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng chọn thông tin muốn sửa", "Thông báo");
                }   
            }
            else
            {
                MessageBox.Show("Vui lòng chọn thông tin muốn sửa", "Thông báo");
            }    
        }

        private void btn_xoa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (dgv_dangnhap.SelectedCells.Count > 0)
            {
                if (!string.IsNullOrWhiteSpace(txt_id.Text) || !string.IsNullOrWhiteSpace(txt_user.Text))
                {
                    try
                    {
                        DialogResult tb = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                        if (tb == DialogResult.OK)
                        {
                            string sql = "delete from PhienDangNhap where Id = '" + txt_id.Text + "' ";
                            int up = my.Update(sql);
                            if (up > 0)
                            {

                                MessageBox.Show("Thông tin được xóa thành công", "Thông báo");
                                loadDL();
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
                    catch
                    {
                        MessageBox.Show("Lỗi xóa thông tin dữ liệu", "Thông báo", MessageBoxButtons.OKCancel);

                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng chọn thông tin muốn xóa", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn thông tin muốn xóa", "Thông báo");
            }
        }

        private void btn_timkiem_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
           
        }

        private void btn_refresh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            loadDL();
        }

        private void dgv_dangnhap_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txt_id.Text = dgv_dangnhap.CurrentRow.Cells[0].Value.ToString();
            txt_user.Text = dgv_dangnhap.CurrentRow.Cells[1].Value.ToString();
            dtp_login.Text = dgv_dangnhap.CurrentRow.Cells[2].Value.ToString();
            dtp_logout.Text = dgv_dangnhap.CurrentRow.Cells[3].Value.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (cb_loai.Text == "")
            {
                MessageBox.Show("Vui lòng chọn khóa tìm kiếm", "Thông báo");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(txt_timkiem.Text))
                {
                    MessageBox.Show("Vui lòng điền thông tin tìm kiếm", "Thông báo");
                }
                else
                {
                    if (cb_loai.Text == "ID")
                    {
                        try
                        {
                            string sql = "select * from PhienDangNhap where Id = '" + txt_timkiem.Text + "' ";
                            DataTable tb = my.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_dangnhap.DataSource = tb;
                                Header();
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm", "Thông báo");
                        }

                        return;
                    }
                    if (cb_loai.Text == "Username")
                    {
                        try
                        {
                            string sql = "select * from PhienDangNhap where Username like '%" + txt_timkiem.Text + "%' ";
                            DataTable tb = my.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_dangnhap.DataSource = tb;
                                Header();
                            }
                        }
                        catch
                        {
                            MessageBox.Show("Lỗi tìm kiếm", "Thông báo");
                        }
                    }
                }
            }
        }
    }
}
