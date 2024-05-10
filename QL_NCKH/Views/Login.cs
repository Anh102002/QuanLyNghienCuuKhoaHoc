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
    public partial class Login : Form
    {
        MyClass my = new MyClass();
        public Login()
        {
           
            InitializeComponent();
            
        }

        private void btn_close_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void chb_pass_CheckedChanged(object sender, EventArgs e)
        {
            if (chb_pass.Checked)
            {
                txt_pass.PasswordChar = (char)0;
            }
            else txt_pass.PasswordChar = '*';
        }

        private void btn_exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void link_quenmatkhau_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            QuenMatKhau f = new QuenMatKhau();
            f.ShowDialog();
            
        }
        
        
        public int login()
        {
            try
            {
                string sql = "select Id from PhienDangNhap";
                DataTable tb = my.DocDL(sql);
                int id2;
                if (tb.Rows.Count > 0)
                {
                    string id = tb.Rows[tb.Rows.Count - 1][0].ToString();
                    int id1 = int.Parse(id);
                    id2 = id1 + 1;


                }
                else
                {
                    id2 = 1;
                }
                
                DateTime time = DateTime.Now;
                string timeFormat = time.ToString("yyyy-MM-dd HH:mm:ss");
                string query = "insert into PhienDangNhap values('" + id2 + "','" + txt_username.Text + "','" + timeFormat + "','')";
                int up = my.Update(query);
                if (up > 0)
                {
                    return id2;
                }
                else
                {
                    MessageBox.Show("Lỗi thêm lịch sử đăng nhập {" + id2 + "} !", "Thông báo");
                    Application.Exit();
                    return 0;

                }



            }
            catch(Exception e)
            {
                MessageBox.Show("Lỗi kiểm tra đăng nhập { "+e.Message+" }!", "Thông báo");
                return 0;
            }
            
        }

        public void checkLogin()
        {
            try
            {
               
                string sql = "SELECT * FROM Account WHERE Username = '" + txt_username.Text + "'  AND Password = '" + txt_pass.Text + "' ";
                
                DataTable tb = my.DocDL(sql);
               
                if (tb.Rows.Count > 0)
                {                    
                    string trangthai = tb.Rows[0]["TrangThai"].ToString();

                    if (trangthai == "Mở")
                    {
                        MainFrm f = new MainFrm();
                        string quyen = tb.Rows[0]["PhanQuyen"].ToString();
                        string Hoten = tb.Rows[0]["HoTen"].ToString();
                        string user = tb.Rows[0]["Username"].ToString();
                        if (quyen =="User")
                        {
                            f.setQuyen("User");
                            f.setHoten (Hoten);
                            f.setUser(txt_username.Text);
                            int id = login();
                            if(id > 0)
                            {
                                f.setId(id);
                                this.Hide();
                                f.ShowDialog();


                                this.Dispose();
                                this.Close();
                                return;
                            }
                            
                        }
                        else
                        {
                           
                            f.setQuyen("Administrators");
                            f.setHoten(Hoten);
                            f.setUser(user);
                            int id = login();
                            if (id > 0)
                            {
                                f.setId(id);
                                this.Hide();
                                f.ShowDialog();
                               

                                this.Dispose();
                                this.Close();
                                return;
                            }
                        }
                        
                        
                    }

                    if (trangthai == "Khóa")
                    {
                        MessageBox.Show("Tài khoản bị khóa","Thông báo", MessageBoxButtons.OKCancel);
                        return;
                    }




                }
                else
                {
                    MessageBox.Show("Tên đăng nhập hoặc mật khẩu sai .Vui lòng nhập lại !!", "Cảnh báo", MessageBoxButtons.OKCancel);
                   
                }
            
            }
            catch
            {
                MessageBox.Show("Lỗi không đăng nhập được ","Thông báo", MessageBoxButtons.OKCancel);

            }



        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            if(string.IsNullOrWhiteSpace(txt_username.Text) && string.IsNullOrWhiteSpace(txt_pass.Text))
            {
                MessageBox.Show("Vui lòng nhập thông tin đăng nhập  ","Thông báo",MessageBoxButtons.OKCancel);
                txt_username.Focus();
                return;
            } 
            
            if(string.IsNullOrWhiteSpace(txt_username.Text))
            {
                MessageBox.Show("Vui lòng nhập Username  ", "Thông báo", MessageBoxButtons.OKCancel);
                txt_username.Focus();
                return;

            }

            if(string.IsNullOrWhiteSpace(txt_pass.Text))
            {
                MessageBox.Show("Vui lòng nhập mật khẩu  ", "Thông báo", MessageBoxButtons.OKCancel);
                txt_pass.Focus();
                return;
            }


            if(!string.IsNullOrWhiteSpace(txt_username.Text) && !string.IsNullOrWhiteSpace(txt_pass.Text))
            {
                checkLogin();
               
            }    

        }

        private void button1_KeyDown(object sender, KeyEventArgs e)
        {
            
        }
    }
}
