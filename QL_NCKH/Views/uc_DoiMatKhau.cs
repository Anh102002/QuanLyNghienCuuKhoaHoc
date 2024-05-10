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
    public partial class uc_DoiMatKhau : DevExpress.XtraEditors.XtraForm
    {
        MyClass my = new MyClass();
        private string user;

        public string getUser()
        {
            return this.user;
        }

        public void setUser(string userN)
        {
            this.user = userN;
        }
        public uc_DoiMatKhau()
        {
            InitializeComponent();
             
        }

        private void btn_thoat_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void uc_DoiMatKhau_Load(object sender, EventArgs e)
        {
            string user = getUser();
            txt_user.Text = user;
            txt_email.Select();
        }

        private void btn_apply_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrWhiteSpace(txt_email.Text) || string.IsNullOrWhiteSpace(txt_pass.Text) || string.IsNullOrWhiteSpace(txt_updatepass.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin!","Thông báo");
            }
            else
            {
                try
                {
                    string sql = "select * from Account where Username = '"+txt_user.Text+"' and Email = '"+txt_email.Text+"' ";
                    DataTable tb = my.DocDL(sql);
                    if(tb.Rows.Count > 0)
                    {
                        if(txt_pass.Text == txt_updatepass.Text)
                        {
                            string query = "update Account set Password = '"+txt_updatepass.Text+"' ";
                             int up = my.Update(query);
                            if(up > 0)
                            {
                                MessageBox.Show("Đổi mật khẩu thành công", "Thông báo");
                                this.Close();
                            }
                            else
                            {
                                MessageBox.Show("Đổi mật khẩu không thành công", "Thông báo");
                            }
                                
                              
                        }
                        else
                        {
                            MessageBox.Show("Vui lòng nhập mật khẩu trùng nhau", "Thông báo");
                        }   
                    }
                    else
                    {
                        MessageBox.Show("Bạn nhập sai Email !", "Thông báo");
                    }   
                }
                catch
                {
                    MessageBox.Show("Lỗi kiểm tra mật khẩu", "Thông báo");

                }
            }    
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if(checkBox1.Checked)
            {
                txt_pass.PasswordChar = (char)0;
                txt_updatepass.PasswordChar = (char)0;
            }
            else
            {
                txt_pass.PasswordChar = '*';
                txt_updatepass.PasswordChar = '*';
            }    
        }
    }
}