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
    public partial class frm_hopdong : DevExpress.XtraEditors.XtraForm
    {


        MyClass my = new MyClass();
        private string makm;
        public frm_hopdong()
        {
            InitializeComponent();
        }


        public string MaKM
        {
            get { return this.makm; }
            set { this.makm = value; }
        }
        public void LoadDL()
        {
            if (MaKM != null)
            {
                string makm = MaKM;
                string sql = "select MaHD,TieuDe,DuongDan from HopDong where MaKM = '" + makm + "' ";
                DataTable tb = my.DocDL(sql);
                dgv_hd.DataSource = tb;
                dgv_hd.Columns[0].HeaderText = "Mã hợp đồng";
                dgv_hd.Columns[1].HeaderText = "Tiêu đề";
                dgv_hd.Columns[2].HeaderText = "Đường dẫn";
                dgv_hd.Columns[2].Width = 300;

            }
        }
        private void frm_hopdong_Load(object sender, EventArgs e)
        {
            LoadDL();
        }

        public bool ktraMa(string ma, string makm)
        {
            string sql = "select * from HopDong where MaKM = '" + makm + "' and MaHD = '" + ma + "' ";
            DataTable tb = my.DocDL(sql);
            if (tb.Rows.Count > 0)
            {
                return false;
            }
            return true;
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

        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (MaKM != null)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_magt.Text) || string.IsNullOrWhiteSpace(txt_tengt.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin hợp đồng", "Thông báo");
                    }
                    else
                    {
                        string makm = MaKM;
                        string ma = txt_magt.Text;
                        if (ktraMa(ma, makm))
                        {
                            string sql = "insert into HopDong values ('" +makm + "','" + txt_magt.Text + "',N'" + txt_tengt.Text + "',N'" + txt_url.Text + "')";
                            int up = my.Update(sql);
                            if (up > 0)
                            {
                                MessageBox.Show("Thêm hợp đồng thành công", "Thông báo");
                                LoadDL();
                                txt_magt.Clear();
                                txt_tengt.Clear();
                                txt_url.Clear();
                            }


                        }
                        else
                        {
                            MessageBox.Show("Trùng mã hợp đồng vui lòng nhập lại !!!", "Thông báo");
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
            if (MaKM != null)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_magt.Text) || string.IsNullOrWhiteSpace(txt_tengt.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin hợp đồng", "Thông báo");
                    }
                    else
                    {
                        string makm = MaKM;
                        string ma = txt_magt.Text;
                        if (!ktraMa(ma, makm))
                        {
                            string sql = "update HopDong set TieuDe=N'" + txt_tengt.Text + "',DuongDan = N'" + txt_url.Text + "' where MaKM ='" + makm + "' and MaHD = '" + ma + "' ";
                            int up = my.Update(sql);
                            if (up > 0)
                            {
                                MessageBox.Show("Sửa hợp đồng thành công", "Thông báo");
                                LoadDL();
                                txt_magt.Clear();
                                txt_tengt.Clear();
                                txt_url.Clear();
                            }


                        }
                        else
                        {
                            MessageBox.Show("Thông tin hợp đồng này không có trên hệ thống !!!", "Thông báo");
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
            if (MaKM != null)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_magt.Text) || string.IsNullOrWhiteSpace(txt_tengt.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin hợp đồng", "Thông báo");
                    }
                    else
                    {
                        string makm = MaKM;
                        string ma = txt_magt.Text;
                        if (!ktraMa(ma, makm))
                        {
                            string sql = "delete from HopDong where MaKM ='" + makm + "' and MaHD = '" + ma + "' ";
                            int up = my.Update(sql);
                            if (up > 0)
                            {
                                MessageBox.Show("Xóa hợp đồng thành công", "Thông báo");
                                LoadDL();
                                txt_magt.Clear();
                                txt_tengt.Clear();
                                txt_url.Clear();
                            }


                        }
                        else
                        {
                            MessageBox.Show("Thông tin hợp đồng này không có trên hệ thống !!!", "Thông báo");
                        }

                    }
                }
                catch
                {
                    MessageBox.Show("Lỗi không xóa được !", "Thông báo");
                }

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

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            this.Close();
        }

        private void dgv_hd_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txt_magt.Text = dgv_hd.CurrentRow.Cells[0].Value.ToString();
            txt_tengt.Text = dgv_hd.CurrentRow.Cells[1].Value.ToString();
            txt_url.Text = dgv_hd.CurrentRow.Cells[2].Value.ToString();
        }
    }
}