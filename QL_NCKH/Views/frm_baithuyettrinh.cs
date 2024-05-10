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
    public partial class frm_baithuyettrinh : DevExpress.XtraEditors.XtraForm
    {
        MyClass my = new MyClass();
        private string madt;

        public frm_baithuyettrinh()
        {
            InitializeComponent();
        }
        public string Madt
        {
            get { return this.madt; }
            set { this.madt = value; }
        }
        public void LoadDL()
        {
            try
            {
                if (Madt != null)
                {
                    string madt = Madt;
                    string sql = "select MaTT,TieuDe,DuongDan from BaiThuyetTrinhCT where MaDoi = '" + madt + "' ";
                    DataTable tb = my.DocDL(sql);
                    dgv_thuyettrinh.DataSource = tb;
                    dgv_thuyettrinh.Columns[0].HeaderText = "Mã bài thuyết trình";
                    dgv_thuyettrinh.Columns[1].HeaderText = "Tiêu đề";
                    dgv_thuyettrinh.Columns[2].HeaderText = "Đường dẫn";
                    dgv_thuyettrinh.Columns[2].Width = 300;

                }
            }
            catch
            {
                MessageBox.Show("Lỗi hiển thị thông tin bài thuyết trình !", "Thông báo");
            }

        }
        private void frm_baithuyettrinh_Load(object sender, EventArgs e)
        {
            LoadDL();
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
        public bool ktraMa(string ma, string madt)
        {
            string sql = "select * from BaiThuyetTrinhCT where MaTT = '" + ma + "' and MaDoi = '" + madt + "' ";
            DataTable tb = my.DocDL(sql);
            if (tb.Rows.Count > 0)
            {
                return false;
            }
            return true;
        }
        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Madt != null)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_mabtt.Text) || string.IsNullOrWhiteSpace(txt_tengt.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                    }
                    else
                    {
                        string madt = Madt;
                        string ma = txt_mabtt.Text;
                        if (ktraMa(ma, madt))
                        {
                            string sql = "insert into BaiThuyetTrinhCT values ('" + txt_mabtt.Text + "','" + madt + "',N'" + txt_tengt.Text + "',N'" + txt_url.Text + "')";
                            int update = my.Update(sql);
                            if (update > 0)
                            {
                                MessageBox.Show("Thêm bài thuyết trình thành công", "Thông báo");
                                LoadDL();
                                txt_mabtt.Clear();
                                txt_tengt.Clear();
                                txt_url.Clear();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Trùng mã , vui lòng nhập lại !!!", "Thông báo");
                        }

                    }
                }
                catch
                {
                    MessageBox.Show("Lỗi không thêm được !", "Lỗi");
                }

            }
        }

        private void btn_sua_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Madt != null)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_mabtt.Text) || string.IsNullOrWhiteSpace(txt_tengt.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                    }
                    else
                    {
                        string madt = Madt;
                        string ma = txt_mabtt.Text;
                        if (!ktraMa(ma, madt))
                        {
                            string sql = "update BaiThuyetTrinhCT set TieuDe=N'" + txt_tengt.Text + "',DuongDan = N'" + txt_url.Text + "' where MaTT ='" + ma + "' and MaDoi= '" + madt + "' ";
                            int update = my.Update(sql);
                            if (update > 0)
                            {
                                MessageBox.Show("Sửa thông tin thành công", "Thông báo");
                                LoadDL();
                                txt_mabtt.Clear();
                                txt_tengt.Clear();
                                txt_url.Clear();
                            }


                        }
                        else
                        {
                            MessageBox.Show("Thông tin này không có trên hệ thống !!!", "Thông báo");
                        }

                    }
                }
                catch
                {
                    MessageBox.Show("Lỗi không sửa được !", "Lỗi");
                }

            }
        }

        private void btn_xoa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (Madt != null)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_mabtt.Text) || string.IsNullOrWhiteSpace(txt_tengt.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                    }
                    else
                    {
                        string madt = Madt;
                        string ma = txt_mabtt.Text;
                        if (!ktraMa(ma, madt))
                        {

                            DialogResult tb = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                            if (tb == DialogResult.OK)
                            {
                                string sql = "delete from BaiThuyetTrinhCT where MaTT ='" + ma + "' and MaDoi= '" + madt + "' ";
                                int update = my.Update(sql);
                                if (update > 0)
                                {
                                    MessageBox.Show("Xóa thông tin thành công", "Thông báo");
                                    LoadDL();
                                    txt_mabtt.Clear();
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
                            MessageBox.Show("Thông tin này không có trên hệ thống !!!", "Thông báo");
                        }

                    }
                }
                catch
                {
                    MessageBox.Show("Lỗi không xóa được !", "Lỗi");
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
                MessageBox.Show("Vui lòng chọn bài thuyết trình !!", "Thông báo");
            }
        }

        private void dgv_thuyettrinh_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_mabtt.Text = dgv_thuyettrinh.CurrentRow.Cells[0].Value.ToString();
                txt_tengt.Text = dgv_thuyettrinh.CurrentRow.Cells[1].Value.ToString();
                txt_url.Text = dgv_thuyettrinh.CurrentRow.Cells[2].Value.ToString();


            }
            catch
            {
                MessageBox.Show("Lỗi lấy dữ liệu bài thuyết trình !", "Thông báo");
            }
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            this.Close();
        }
    }
}