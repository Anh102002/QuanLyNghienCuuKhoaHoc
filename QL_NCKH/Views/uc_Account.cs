using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QL_NCKH
{
    public partial class uc_Account : DevExpress.XtraEditors.XtraUserControl
    {
        MyClass myClass;
        public uc_Account()
        {
            InitializeComponent();
            myClass = new MyClass();
        }
        public void Header()
        {
            
            dgv_account.Columns[0].HeaderText = "Username";
            dgv_account.Columns[1].HeaderText = "Password";
            dgv_account.Columns[2].HeaderText = "Họ Tên";
            dgv_account.Columns[2].Width = 150;
            dgv_account.Columns[3].HeaderText = "Email";
            dgv_account.Columns[3].Width = 150;
            dgv_account.Columns[4].HeaderText = "Phân Quyền";
            dgv_account.Columns[5].HeaderText = "Trạng Thái";
        }
        public void loadDL()
        {
            try
            {
                string sql = "select * from Account ";
                DataTable tb = myClass.DocDL(sql);
                if(tb.Rows.Count > 0)
                {
                    dgv_account.DataSource = tb;
                    Header();



                }
               
            }
            catch
            {
                MessageBox.Show("Lỗi thực hiện lấy dữ liệu tài khoản","Thông báo");
            }
        }
        private void uc_Account_Load(object sender, EventArgs e)
        {
            loadDL();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dgv_account_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           


        }
        public bool ktraMa()
        {
            try
            {
                string sql = "select * from Account where Username = '" + txt_user.Text + "' ";
                DataTable tb = myClass.DocDL(sql);
                if(tb.Rows.Count > 0)
                {
                    return false;
                }    
            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra trùng tài khoản", "Thông báo");

            }
            return true;
        }

        public bool ktraNull()
        {
            if (string.IsNullOrWhiteSpace(txt_user.Text) || string.IsNullOrWhiteSpace(txt_pass.Text) || string.IsNullOrWhiteSpace(txt_hoten.Text)
                || string.IsNullOrWhiteSpace(txt_email.Text) || string.IsNullOrWhiteSpace(cb_quyen.Text) || string.IsNullOrWhiteSpace(cb_trangthai.Text))
                return false;
            return true;
        }
        public bool IsValidEmail(string email)
        {
            string pattern = @"^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$";
            Regex regex = new Regex(pattern);
            return regex.IsMatch(email);
        }
        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                if (ktraNull())
                {
                    if (IsValidEmail(txt_email.Text))
                    {
                        if (ktraMa())
                        {
                            string sql = "insert into Account values('" + txt_user.Text + "','" + txt_pass.Text + "',N'" + txt_hoten.Text + "','" + txt_email.Text + "',N'" + cb_quyen.Text + "',N'" + cb_trangthai.Text + "')";
                           int up = myClass.Update(sql);
                            if(up  > 0 )
                            {
                                MessageBox.Show("Thêm thông tin thành công", "Thông báo");
                                loadDL();
                                txt_user.Clear();
                                txt_hoten.Clear();
                                txt_email.Clear();
                                txt_pass.Clear();
                                txt_timkiem.Clear();
                                cb_quyen.SelectedIndex = -1;
                                cb_loai.SelectedIndex = -1;
                                cb_trangthai.SelectedIndex = -1;
                            }
                                

                            
                           
                        }
                        else
                        {
                            MessageBox.Show("Username này đã có trong hệ thống", "Thông báo");

                        }
                    }
                    else
                    {
                        MessageBox.Show("Vui lòng nhập đúng định dạng Email", "Thông báo");

                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin!", "Thông báo");

                }
            }
            catch
            {
                MessageBox.Show("Lỗi thêm dữ liệu", "Thông báo");

            }
        }

        private void btn_sua_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                if(dgv_account.SelectedCells.Count > 0)
                {
                    if (ktraNull())
                    {
                        if (IsValidEmail(txt_email.Text))
                        {
                            if (!ktraMa())
                            {
                                string sql = "update Account set Password='" + txt_pass.Text + "' ,HoTen=N'" + txt_hoten.Text + "' , Email='" + txt_email.Text + "' ,PhanQuyen = N'" + cb_quyen.Text + "' ,TrangThai= N'" + cb_trangthai.Text + "' where Username ='" + txt_user.Text + "' ";
                                int up = myClass.Update(sql);
                                if (up > 0)
                                {
                                    MessageBox.Show("Sửa thông tin thành công", "Thông báo");
                                    loadDL();
                                    txt_user.Clear();
                                    txt_hoten.Clear();
                                    txt_email.Clear();
                                    txt_pass.Clear();
                                    txt_timkiem.Clear();
                                    cb_quyen.SelectedIndex = -1;
                                    cb_loai.SelectedIndex = -1;
                                    cb_trangthai.SelectedIndex = -1;
                                }

                            }
                            else
                            {
                                MessageBox.Show("Username này không có trong hệ thống", "Thông báo");

                            }
                        }
                        else
                        {
                            MessageBox.Show("Vui lòng nhập đúng định dạng Email", "Thông báo");

                        }
                    }
                    else
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin!", "Thông báo");

                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng chọn tài khoản cần sửa", "Thông báo");

                }
                
            }
            catch
            {
                MessageBox.Show("Lỗi sửa dữ liệu", "Thông báo");

            }
        }

        private void btn_xoa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                if (dgv_account.SelectedCells.Count > 0)
                {
                    if(ktraNull())
                    {
                        if (IsValidEmail(txt_email.Text))
                        {
                            if (!ktraMa())
                            {

                                DialogResult tb = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                                if (tb == DialogResult.OK)
                                {
                                    string pdn = "delete from PhienDangNhap where Username = '" + txt_user.Text + "' ";
                                    int upPDN = myClass.Update(pdn);
                                    if (upPDN >= 0)
                                    {
                                        string sql = "delete from Account where Username = '" + txt_user.Text + "' ";
                                        int up = myClass.Update(sql);
                                        if (up > 0)
                                        {
                                            MessageBox.Show("Xóa thông tin thành công", "Thông báo");
                                            loadDL();
                                            txt_user.Clear();
                                            txt_hoten.Clear();
                                            txt_email.Clear();
                                            txt_pass.Clear();
                                            txt_timkiem.Clear();
                                            cb_quyen.SelectedIndex = -1;
                                            cb_loai.SelectedIndex = -1;
                                            cb_trangthai.SelectedIndex = -1;
                                        }
                                    }
                                }
                                else
                                {

                                }
                                   




                                


                            }
                            else
                            {
                                MessageBox.Show("Username này không có trong hệ thống", "Thông báo");

                            }
                        }
                        else
                        {
                            MessageBox.Show("Vui lòng nhập đúng định dạng Email", "Thông báo");

                        }
                    }
                    else
                    {
                        MessageBox.Show("Vui lòng chọn tài khoản muốn xóa!", "Thông báo");

                    }
                }
                    else
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin", "Thông báo");

                    }
       
            }
            catch
            {
                MessageBox.Show("Lỗi xóa dữ liệu", "Thông báo");

            }
        }

        private void btn_timkiem_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            txt_user.Clear();
            txt_pass.Clear();
            txt_hoten.Clear();
            txt_email.Clear();
            cb_quyen.SelectedIndex = -1 ;
            cb_trangthai.SelectedIndex = -1 ;
           
            if (cb_loai.Text == "")
            {
                MessageBox.Show("Vui lòng chọn khóa tìm kiếm ! ", "Thông báo");
            }
            else
            {
                if (txt_timkiem.Text != "")
                {
                    if (cb_loai.Text == "Username")
                    {
                        try
                        {
                            string sql = "select * from Account where Username like '%" + txt_timkiem.Text + "%' ";
                            DataTable tb = myClass.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_account.DataSource = tb;
                                Header();

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
                            string sql = "select * from Account where HoTen like N'%" + txt_timkiem.Text + "%' ";
                            DataTable tb = myClass.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_account.DataSource = tb;
                                Header();

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


                    if (cb_loai.Text == "Email")
                    {
                        try
                        {
                            string sql = "select * from Account where Email like '%" + txt_timkiem.Text + "%' ";
                            DataTable tb = myClass.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_account.DataSource = tb;
                                Header();

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

        private void dgv_account_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_user.Text = dgv_account.CurrentRow.Cells[0].Value.ToString();
                txt_pass.Text = dgv_account.CurrentRow.Cells[1].Value.ToString();
                txt_hoten.Text = dgv_account.CurrentRow.Cells[2].Value.ToString();
                txt_email.Text = dgv_account.CurrentRow.Cells[3].Value.ToString();
                cb_quyen.Text = dgv_account.CurrentRow.Cells[4].Value.ToString();
                cb_trangthai.Text = dgv_account.CurrentRow.Cells[5].Value.ToString();
            }
            catch
            {
                MessageBox.Show("Lỗi thực hiện lấy dữ liệu từ bảng", "Thông báo");

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            txt_user.Clear();
            txt_pass.Clear();
            txt_hoten.Clear();
            txt_email.Clear();
            cb_quyen.SelectedIndex = -1;
            cb_trangthai.SelectedIndex = -1;

            if (cb_loai.Text == "")
            {
                MessageBox.Show("Vui lòng chọn khóa tìm kiếm ! ", "Thông báo");
            }
            else
            {
                if (txt_timkiem.Text != "")
                {
                    if (cb_loai.Text == "Username")
                    {
                        try
                        {
                            string sql = "select * from Account where Username like '%" + txt_timkiem.Text + "%' ";
                            DataTable tb = myClass.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_account.DataSource = tb;
                                Header();

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
                            string sql = "select * from Account where HoTen like N'%" + txt_timkiem.Text + "%' ";
                            DataTable tb = myClass.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_account.DataSource = tb;
                                Header();

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


                    if (cb_loai.Text == "Email")
                    {
                        try
                        {
                            string sql = "select * from Account where Email like '%" + txt_timkiem.Text + "%' ";
                            DataTable tb = myClass.DocDL(sql);
                            if (tb.Rows.Count > 0)
                            {
                                dgv_account.DataSource = tb;
                                Header();

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
