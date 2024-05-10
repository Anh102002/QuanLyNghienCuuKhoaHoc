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
using System.Diagnostics;

namespace QL_NCKH
{
    public partial class uc_giaytoluutru : DevExpress.XtraEditors.XtraUserControl
    {
        public uc_giaytoluutru()
        {
            InitializeComponent();
        }
        MyClass my = new MyClass();
        public void LoadDL()
        {
            try
            {
                string sql = "select * from GiayToVanBan  ";
                DataTable tb = my.DocDL(sql);
                dgv_giayto.DataSource = tb;
                dgv_giayto.Columns[0].HeaderText = "Mã giấy tờ";
                dgv_giayto.Columns[1].HeaderText = "Tiêu đề";
                dgv_giayto.Columns[2].HeaderText = "Đường dẫn";
                dgv_giayto.Columns[3].HeaderText = "Đối tượng";
                dgv_giayto.Columns[2].Width = 300;
            }catch
            {
                MessageBox.Show("Lỗi lấy dữ liệu giấy tờ","Lỗi");
            }
                

            
        }
        private void uc_giaytoluutru_Load(object sender, EventArgs e)
        {

            LoadDL();
        }
        public bool ktraMa()
        {
            string sql = "select * from GiayToVanBan where MaGT = '" + txt_magt.Text+"' ";
            DataTable tb = my.DocDL(sql);
            if (tb.Rows.Count > 0)
            {
                return false;
            }
            return true;
        }
        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txt_magt.Text) || string.IsNullOrWhiteSpace(txt_tengt.Text)
                    || string.IsNullOrWhiteSpace(txt_url.Text) || string.IsNullOrWhiteSpace(cbo_luutru.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin giấy tờ", "Thông báo");
                }
                else
                {
                    
                    
                    if (ktraMa())
                    {
                        string sql = "insert into GiayToVanBan values ('" + txt_magt.Text + "',N'" +txt_tengt.Text  + "',N'" + txt_url.Text + "',N'" + cbo_luutru.Text + "')";
                        int up = my.Update(sql);
                        if (up > 0)
                        {
                            MessageBox.Show("Thêm giấy tờ thành công", "Thông báo");
                            LoadDL();                           
                            txt_tengt.Clear();
                            txt_magt.Clear();
                            txt_url.Clear();
                            cbo_luutru.SelectedIndex = -1;
                        }



                    }
                    else
                    {
                        MessageBox.Show("Trùng mã giấy tờ vui lòng nhập lại !!!", "Thông báo");
                    }

                }
            }
            catch
            {
                MessageBox.Show("Lỗi không thêm được !", "Thông báo");
            }
        }

        private void btn_sua_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txt_magt.Text) || string.IsNullOrWhiteSpace(txt_tengt.Text)
                    || string.IsNullOrWhiteSpace(txt_url.Text) || string.IsNullOrWhiteSpace(cbo_luutru.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin giấy tờ", "Thông báo");
                }
                else
                {


                    if (!ktraMa())
                    {
                        string sql = "update GiayToVanBan set TieuDe= N'" + txt_tengt.Text + "',DuongDan=N'" + txt_url.Text + "',DoiTuong=N'" + cbo_luutru.Text + "'  where MaGT = '"+txt_magt.Text+"' ";
                        int up = my.Update(sql);
                        if (up > 0)
                        {
                            MessageBox.Show("Sửa giấy tờ thành công", "Thông báo");
                            LoadDL();
                            txt_tengt.Clear();
                            txt_magt.Clear();
                            txt_url.Clear();
                            cbo_luutru.SelectedIndex = -1;
                        }



                    }
                    else
                    {
                        MessageBox.Show("không có giấy tờ này !!!", "Thông báo");
                    }

                }
            }
            catch
            {
                MessageBox.Show("Lỗi không sửa được !", "Thông báo");
            }
        }

        private void btn_xoa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txt_magt.Text )|| string.IsNullOrWhiteSpace(txt_tengt.Text)
                    || string.IsNullOrWhiteSpace(txt_url.Text) || string.IsNullOrWhiteSpace(cbo_luutru.Text))
                {
                    MessageBox.Show("Vui lòng nhập đầy đủ thông tin giấy tờ", "Thông báo");
                }
                else
                {


                    if (!ktraMa())
                    {
                        DialogResult tb = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                        if (tb == DialogResult.OK)
                        {

                            string sql = "delete from GiayToVanBan where MaGT = '" + txt_magt.Text + "' ";
                            int up = my.Update(sql);
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa giấy tờ thành công", "Thông báo");
                                LoadDL();
                                txt_magt.Clear();
                                txt_tengt.Clear();
                                txt_url.Clear();
                                cbo_luutru.SelectedIndex = -1;
                            }
                        }
                        else
                        {

                        }




                    }
                    else
                    {
                        MessageBox.Show("không có giấy tờ này !!!", "Thông báo");
                    }

                }
            }
            catch
            {
                MessageBox.Show("Lỗi không xóa được !", "Thông báo");
            }
        }

        private void btn_openfile_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string filePath = txt_url.Text;
            if (!string.IsNullOrEmpty(filePath))
            {

                Process.Start(filePath);
            }
            else
            {
                MessageBox.Show("Vui lòng chọn tập tin văn bản !!", "Thông báo");
            }
        }

        private void btn__choose_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();


            openFileDialog.Title = "Chọn tập tin";
            openFileDialog.Filter = "Các loại tập tin (*.txt;*.csv;*.docx)|*.txt;*.csv;*.docx|Tất cả các tập tin (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {

                string selectedFilePath = openFileDialog.FileName;
                txt_url.Text = selectedFilePath;
            }
        }

        private void dgv_giayto_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            
                txt_magt.Text= dgv_giayto.CurrentRow.Cells[0].Value.ToString();
            
                
                txt_tengt.Text = dgv_giayto.CurrentRow.Cells[1].Value.ToString();
                txt_url.Text = dgv_giayto.CurrentRow.Cells[2].Value.ToString();
                cbo_luutru.Text = dgv_giayto.CurrentRow.Cells[3].Value.ToString();
            

                
            
            
        }
    }
}
