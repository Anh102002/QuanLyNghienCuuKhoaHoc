using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Diagnostics;

namespace QL_NCKH
{
    public partial class frm_giaytoSv : DevExpress.XtraEditors.XtraForm
    {
        MyClass my = new MyClass();
        private string magv;
        public frm_giaytoSv()
        {
            InitializeComponent();
        }
        public string Magv
        {
            get { return this.magv; }
            set { this.magv = value; }
        }
        public void LoadDL()
        {
            if (Magv != null)
            {
                string magv = Magv;
                string sql = "select MaGT,TenGT,Duongdan from GiayToGV where MaGV = '" + magv + "' ";
                DataTable tb = my.DocDL(sql);
                dgv_giayto.DataSource = tb;
                dgv_giayto.Columns[0].HeaderText = "Mã giấy tờ";
                dgv_giayto.Columns[1].HeaderText = "Tên giấy tờ";
                dgv_giayto.Columns[2].HeaderText = "Đường dẫn";
                dgv_giayto.Columns[2].Width = 300;

            }
        }
        private void frm_giayto_Load(object sender, EventArgs e)
        {


            LoadDL();

        }

        private void dgv_giayto_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }
        public bool ktraMa(string ma, string magv)
        {
            string sql = "select * from GiayToGV where MaGT = '"+ma+"' and MaGV = '"+magv+"' ";
            DataTable tb = my.DocDL(sql);
            if(tb.Rows.Count > 0)
            {
                return false;
            }
            return true;
        }
        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if(Magv != null)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_magt.Text) || string.IsNullOrWhiteSpace(txt_tengt.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin giấy tờ", "Thông báo");
                    }
                    else
                    {
                        string magv = Magv;
                        string ma = txt_magt.Text;
                        if (ktraMa(ma,magv))
                        {
                            string sql = "insert into GiayToGV values ('" + txt_magt.Text + "','" + magv + "',N'" + txt_tengt.Text + "',N'" + txt_url.Text + "')";
                            int up = my.Update(sql);
                            if(up > 0)
                            {
                                MessageBox.Show("Thêm giấy tờ giảng viên thành công", "Thông báo");
                                LoadDL();
                                txt_magt.Clear();
                                txt_tengt.Clear();
                                txt_url.Clear();
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

        private void btn_sua_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Magv != null)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_magt.Text) || string.IsNullOrWhiteSpace(txt_tengt.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin giấy tờ", "Thông báo");
                    }
                    else
                    {
                        string magv = Magv;
                        string ma = txt_magt.Text;
                        if (!ktraMa(ma, magv))
                        {
                            string sql = "update GiayToGV set TenGT=N'"+txt_tengt.Text+"',DuongDan = N'"+txt_url.Text+"' where MaGT ='"+ma+"' and MaGV= '"+magv+"' ";
                           int up = my.Update(sql);
                            if(up > 0)
                            {
                                MessageBox.Show("Sửa giấy tờ giảng viên thành công", "Thông báo");
                                LoadDL();
                                txt_magt.Clear();
                                txt_tengt.Clear();
                                txt_url.Clear();
                            }
                               
                            
                        }
                        else
                        {
                            MessageBox.Show("Thông tin giấy tờ này không có trên hệ thống !!!", "Thông báo");
                        }

                    }
                }
                catch
                {
                    MessageBox.Show("Lỗi không sửa được !", "Thông báo");
                }

            }
        }

        private void btn_xoa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Magv != null)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_magt.Text) || string.IsNullOrWhiteSpace(txt_tengt.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin giấy tờ", "Thông báo");
                    }
                    else
                    {
                        string magv = Magv;
                        string ma = txt_magt.Text;
                        if (!ktraMa(ma, magv))
                        {

                            DialogResult tb = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                            if (tb == DialogResult.OK)
                            {
                                string sql = "delete from GiayToGV where MaGT = '" + ma + "' and MaGV = '" + magv + "' ";
                                int up = my.Update(sql);
                                if (up > 0)
                                {
                                    MessageBox.Show("Xóa giấy tờ giảng viên thành công", "Thông báo");
                                    LoadDL();
                                    txt_magt.Clear();
                                    txt_tengt.Clear();
                                    txt_url.Clear();
                                }
                            }
                            else
                            {

                            }
                             
                               
                            
                        }
                        else
                        {
                            MessageBox.Show("Thông tin giấy tờ này không có trên hệ thống !!!", "Thông báo");
                        }

                    }
                }
                catch
                {
                    MessageBox.Show("Lỗi không xóa được !", "Thông báo");
                }

            }
        }

        private void btn_timkiem_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

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
                MessageBox.Show("Vui lòng chọn tập tin văn bản !!","Thông báo");
            }
        }

        private void dgv_giayto_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txt_magt.Text = dgv_giayto.CurrentRow.Cells[0].Value.ToString();
            txt_tengt.Text = dgv_giayto.CurrentRow.Cells[1].Value.ToString();
            txt_url.Text = dgv_giayto.CurrentRow.Cells[2].Value.ToString();
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            this.Close();
        }
    }
}