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
    public partial class QuenMatKhau : DevExpress.XtraEditors.XtraForm
    {
        MyClass myClass;

        public QuenMatKhau()
        {
            InitializeComponent();
            myClass = new MyClass();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txt_user.Text) || string.IsNullOrWhiteSpace(txt_email.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin !", "Thông báo");
                }
                else
                {
                    string sql = "select * from Account where Username = '" + txt_user.Text + "' AND Email = '" + txt_email.Text + "' ";
                    DataTable tb = myClass.DocDL(sql);
                    if (tb.Rows.Count > 0)
                    {
                        txt_resetMk.Text = tb.Rows[0]["Password"].ToString();
                    }
                    else
                    {
                        MessageBox.Show("Thông tin người dùng không có !", "Thông báo");

                    }
                }
            }
            catch
            {
                MessageBox.Show("Lỗi không khôi phục mật khẩu được", "Thông báo");

            }
        }

        private void btn_thoat_Click(object sender, EventArgs e)
        {
            this.Close();

        }
    }
}