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
    public partial class frm_filebaocao : DevExpress.XtraEditors.XtraForm
    {
        MyClass my = new MyClass();
        private string madt;
        public frm_filebaocao()
        {
            InitializeComponent();
        }
        public string Madetai
        {
            get{ return this.madt; }
            set{  this.madt = value ; }
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
        public void LoadDL()
        {
            try
            {
                if (Madetai != null)
                {
                    string madt = Madetai;
                    string sql = "select MaBC,TieuDe,UrlFile from BanBaoCao where MaDeTai = '" + madt + "' ";
                    DataTable tb = my.DocDL(sql);
                    dgv_giayto.DataSource = tb;
                    dgv_giayto.Columns[0].HeaderText = "Mã báo cáo";
                    dgv_giayto.Columns[1].HeaderText = "Tiêu đề";
                    dgv_giayto.Columns[2].HeaderText = "Đường dẫn";
                    dgv_giayto.Columns[2].Width = 300;

                }
            }
            catch
            {
                MessageBox.Show("Lỗi hiển thị thông tin bài báo cáo !","Thông báo");
            }
            
        }
        public bool ktraMa(string ma, string madt)
        {
            string sql = "select * from BanBaoCao where MaBC = '" + ma + "' and MaDeTai = '" + madt + "' ";
            DataTable tb = my.DocDL(sql);
            if (tb.Rows.Count > 0)
            {
                return false;
            }
            return true;
        }
        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Madetai != null)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_magt.Text) || string.IsNullOrWhiteSpace(txt_tengt.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin bài báo cáo", "Thông báo");
                    }
                    else
                    {
                        string madt = Madetai;
                        string ma = txt_magt.Text;
                        if (ktraMa(ma, madt))
                        {
                            string sql = "insert into BanBaoCao values ('" + txt_magt.Text + "','" + madt + "',N'" + txt_tengt.Text + "',N'" + txt_url.Text + "')";
                            int update = my.Update(sql);
                            if(update > 0)
                            {
                                MessageBox.Show("Thêm bài báo cáo thành công", "Thông báo");
                                LoadDL();
                                txt_magt.Clear();
                                txt_tengt.Clear();
                                txt_url.Clear();
                            }                                                            
                        }
                        else
                        {
                            MessageBox.Show("Trùng mã báo cáo vui lòng nhập lại !!!", "Thông báo");
                        }

                    }
                }
                catch
                {
                    MessageBox.Show("Lỗi không thêm được !", "Thông báo");
                }

            }
        }

        private void btn_sua_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Madetai != null)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_magt.Text) || string.IsNullOrWhiteSpace(txt_tengt.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin báo cáo", "Thông báo");
                    }
                    else
                    {
                        string madt = Madetai;
                        string ma = txt_magt.Text;
                        if (!ktraMa(ma, madt))
                        {
                            string sql = "update BanBaoCao set TieuDe=N'" + txt_tengt.Text + "',UrlFile = N'" + txt_url.Text + "' where MaBC ='" + ma + "' and MaDeTai= '" + madt + "' ";
                            int update = my.Update(sql);
                            if(update > 0)
                            {
                                MessageBox.Show("Sửa bài báo cáo thành công", "Thông báo");
                                LoadDL();
                                txt_magt.Clear();
                                txt_tengt.Clear();
                                txt_url.Clear();
                            }
                                
                            
                        }
                        else
                        {
                            MessageBox.Show("Thông tin báo cáo này không có trên hệ thống !!!", "Thông báo");
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
            if (Madetai != null)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_magt.Text) || string.IsNullOrWhiteSpace(txt_tengt.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin báo cáo", "Thông báo");
                    }
                    else
                    {
                        string madt = Madetai;
                        string ma = txt_magt.Text;
                        if (!ktraMa(ma, madt))
                        {

                            DialogResult tb = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                            if (tb == DialogResult.OK)
                            {
                                string sql = "delete from BanBaoCao  where MaBC ='" + ma + "' and MaDeTai= '" + madt + "' ";
                                int update = my.Update(sql);
                                if (update > 0)
                                {
                                    MessageBox.Show("Xáo bài báo cáo thành công", "Thông báo");
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
                            MessageBox.Show("Thông tin báo cáo này không có trên hệ thống !!!", "Thông báo");
                        }

                    }
                }
                catch
                {
                    MessageBox.Show("Lỗi không xóa được !", "Thông báo");
                }

            }
        }

        private void frm_filebaocao_Load(object sender, EventArgs e)
        {
            LoadDL();
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
                MessageBox.Show("Vui lòng chọn bài báo cáo !!", "Thông báo");
            }
        }

        private void dgv_giayto_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //try
            //{
            //    txt_magt.Text = dgv_giayto.CurrentRow.Cells[0].Value.ToString();
            //    txt_tengt.Text = dgv_giayto.CurrentRow.Cells[1].Value.ToString();
            //    txt_url.Text = dgv_giayto.CurrentRow.Cells[2].Value.ToString();
                

            //}
            //catch
            //{
            //    MessageBox.Show("Lỗi lấy dữ liệu lên textbox !", "Thông báo");
            //}
        }

        private void dgv_giayto_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_magt.Text = dgv_giayto.CurrentRow.Cells[0].Value.ToString();
                txt_tengt.Text = dgv_giayto.CurrentRow.Cells[1].Value.ToString();
                txt_url.Text = dgv_giayto.CurrentRow.Cells[2].Value.ToString();


            }
            catch
            {
                MessageBox.Show("Lỗi lấy dữ liệu lên textbox !", "Thông báo");
            }
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            this.Close();
        }
    }
}