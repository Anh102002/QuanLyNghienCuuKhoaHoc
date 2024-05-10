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
using System.Data.SqlClient;
using System.Diagnostics;

namespace QL_NCKH
{
    public partial class frm_baibaihoithao : DevExpress.XtraEditors.XtraForm
    {

        MyClass my = new MyClass();
        private string maht;
        public frm_baibaihoithao()
        {
            InitializeComponent();
        }
        public string Maht
        {
            get { return this.maht; }
            set { this.maht = value; }
        }
        public void LoadDL()
        {
            try
            {
                if (Maht != null)
                {
                    string maht = Maht;
                    string sql = "select MaBB,TacGia,TenBB,DuongDan from BaibaoHT where MaHT = '" + maht + "' ";
                    DataTable tb = my.DocDL(sql);
                    dgv_thuyettrinh.DataSource = tb;
                    dgv_thuyettrinh.Columns[0].HeaderText = "Mã bài báo";
                    dgv_thuyettrinh.Columns[1].HeaderText = "Tác giả";
                    dgv_thuyettrinh.Columns[2].HeaderText = "Tên bài báo";
                    dgv_thuyettrinh.Columns[2].Width = 300;
                    dgv_thuyettrinh.Columns[3].HeaderText = "Đường dẫn";

                }
            }
            catch
            {
                MessageBox.Show("Lỗi hiển thị thông tin bài báo !", "Thông báo");
            }

        }
        private void frm_baibaihoithao_Load(object sender, EventArgs e)
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
            string sql = "select * from BaibaoHT where MaBB = '" + ma + "' and MaHT = '" + madt + "' ";
            DataTable tb = my.DocDL(sql);
            if (tb.Rows.Count > 0)
            {
                return false;
            }
            return true;
        }
        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(Maht))
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_mabb.Text) || string.IsNullOrWhiteSpace(txt_tentg.Text) 
                        || string.IsNullOrWhiteSpace(txt_tenbb.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                    }
                    else
                    {
                        string maht = Maht;
                        string ma = txt_mabb.Text;
                        if (ktraMa(ma, maht))
                        {
                            string sql = "insert into BaibaoHT values (@Mabb,@Maht,@Tentg,@Tenbb,@Url)";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Mabb", txt_mabb.Text);
                            command.Parameters.AddWithValue("@Maht", maht);
                            command.Parameters.AddWithValue("@Tentg", txt_tentg.Text);
                            command.Parameters.AddWithValue("@Tenbb", txt_tenbb.Text);
                            command.Parameters.AddWithValue("@Url", txt_url.Text);
                            int update = command.ExecuteNonQuery();
                            if (update > 0)
                            {
                                MessageBox.Show("Thêm bài thành công", "Thông báo");
                                LoadDL();
                                txt_mabb.Clear();
                                txt_tentg.Clear();
                                txt_url.Clear();
                                txt_tenbb.Clear();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Trùng mã , vui lòng nhập lại !!!", "Thông báo");
                        }

                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Lỗi không thêm được {"+ex.Message+"}!", "Lỗi");
                }

            }
            else
            {
                MessageBox.Show("Vui lòng chọn hội thảo !!!", "Thông báo");
            }
        }

        private void btn_sua_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(Maht))
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_mabb.Text) || string.IsNullOrWhiteSpace(txt_tentg.Text)
                        || string.IsNullOrWhiteSpace(txt_tenbb.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                    }
                    else
                    {
                        string maht = Maht;
                        string ma = txt_mabb.Text;
                        if (!ktraMa(ma, maht))
                        {
                            string sql = "update BaibaoHT set TacGia=@Tentg,TenBB=@Tenbb,DuongDan=@Url where MaBB =@Mabb and MaHT=@Maht ";
                            SqlCommand command = my.SqlCommand(sql);
                            command.Parameters.AddWithValue("@Mabb", txt_mabb.Text);
                            command.Parameters.AddWithValue("@Maht", maht);
                            command.Parameters.AddWithValue("@Tentg", txt_tentg.Text);
                            command.Parameters.AddWithValue("@Tenbb", txt_tenbb.Text);
                            command.Parameters.AddWithValue("@Url", txt_url.Text);
                            int update = command.ExecuteNonQuery();
                            if (update > 0)
                            {
                                MessageBox.Show("Sửa bài thành công", "Thông báo");
                                LoadDL();
                                txt_mabb.Clear();
                                txt_tentg.Clear();
                                txt_url.Clear();
                                txt_tenbb.Clear();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Không có bài báo này !!!", "Thông báo");
                        }

                    }
                }
                catch
                {
                    MessageBox.Show("Lỗi không sửa được !", "Lỗi");
                }

            }
            else
            {
                MessageBox.Show("Vui lòng chọn hội thảo !!!", "Thông báo");
            }
        }

        private void btn_xoa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(Maht))
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(txt_mabb.Text) || string.IsNullOrWhiteSpace(txt_tentg.Text)
                        || string.IsNullOrWhiteSpace(txt_tenbb.Text) || string.IsNullOrWhiteSpace(txt_url.Text))
                    {
                        MessageBox.Show("Vui lòng nhập đầy đủ thông tin ", "Thông báo");
                    }
                    else
                    {
                        string maht = Maht;
                        string ma = txt_mabb.Text;
                        if (!ktraMa(ma, maht))
                        {
                            DialogResult tb  = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                            if(tb == DialogResult.OK)
                            {
                                string sql = "delete from BaibaoHT  where MaBB =@Mabb and MaHT=@Maht ";
                                SqlCommand command = my.SqlCommand(sql);
                                command.Parameters.AddWithValue("@Mabb", txt_mabb.Text);
                                command.Parameters.AddWithValue("@Maht", maht);
                                //command.Parameters.AddWithValue("@Tentg", txt_tentg.Text);
                                //command.Parameters.AddWithValue("@Tenbb", txt_tenbb.Text);
                                //command.Parameters.AddWithValue("@Url", txt_url.Text);
                                int update = command.ExecuteNonQuery();
                                if (update > 0)
                                {
                                    MessageBox.Show("Xóa bài thành công", "Thông báo");
                                    LoadDL();
                                    txt_mabb.Clear();
                                    txt_tentg.Clear();
                                    txt_url.Clear();
                                    txt_tenbb.Clear();
                                }
                                else
                                {

                                }
                            }

                            
                        }
                        else
                        {
                            MessageBox.Show("Không có bài báo này !!!", "Thông báo");
                        }

                    }
                }
                catch
                {
                    MessageBox.Show("Lỗi không xóa được !", "Lỗi");
                }

            }
            else
            {
                MessageBox.Show("Vui lòng chọn hội thảo !!!", "Thông báo");
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
                MessageBox.Show("Vui lòng chọn bài báo !!", "Thông báo");
            }
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            this.Close();
        }

        private void dgv_thuyettrinh_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_mabb.Text = dgv_thuyettrinh.CurrentRow.Cells[0].Value.ToString();
                txt_tentg.Text = dgv_thuyettrinh.CurrentRow.Cells[1].Value.ToString();
                txt_tenbb.Text = dgv_thuyettrinh.CurrentRow.Cells[2].Value.ToString();
                txt_url.Text = dgv_thuyettrinh.CurrentRow.Cells[3].Value.ToString();


            }
            catch
            {
                MessageBox.Show("Lỗi lấy dữ liệu bài báo !", "Thông báo");
            }
        }
    }
}