using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Excel = Microsoft.Office.Interop.Excel;
namespace QL_NCKH
{
    public partial class uc_baibaoquocte : DevExpress.XtraEditors.XtraUserControl
    {
        MyClass my = new MyClass();
        private List<string> productList;
        const string LoaiBaiBao = "Quốc tế";
        private string mabb;
        public uc_baibaoquocte()
        {
            InitializeComponent();
        }
        public string Mabb
        {
            get { return this.mabb; }
            set { this.mabb = value; }
        }
        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
        public void LoadDL()
        {
            try
            {
                string sql = "select MaBB,TenBaiBao,TenTapChi,SoVaThoiGianXB,ChiSoISN,NamXB,DiemCongTrinh,GhiChu,TacGia from BaiBao where LoaiBaiBao =N'" + LoaiBaiBao+"' ";
                DataTable tb = my.DocDL(sql);
                
                    dgv_baibao.DataSource = tb;
                    dgv_baibao.Columns[0].HeaderText = "Mã bài báo";
                    dgv_baibao.Columns[1].HeaderText = "Tên bài báo";
                    dgv_baibao.Columns[1].Width = 250;
                    dgv_baibao.Columns[2].HeaderText = "Tên tạp chí";
                    dgv_baibao.Columns[2].Width = 250;
                    dgv_baibao.Columns[3].HeaderText = "Số vào Thời gian xuất bản";
                    dgv_baibao.Columns[3].Width = 150;
                    dgv_baibao.Columns[4].HeaderText = "Chỉ số ISN";
                    dgv_baibao.Columns[5].HeaderText = "Năm xuất bản";
                    dgv_baibao.Columns[6].HeaderText = "Điểm công trình";
                    dgv_baibao.Columns[7].HeaderText = "Ghi chú";
                    dgv_baibao.Columns[8].HeaderText = "Tác giả";
                    dgv_baibao.Columns[8].Width = 300;



            }
            catch
            {
                MessageBox.Show("Lỗi lấy dữ liệu vào danh sách bài báo", "Thông báo");
            }
        }

        public void loadDLTG(string ma)
        {
            try
            {

                string sql = "select GiangVien.MaGV,GiangVien.HoTen,TacGiaBaiBao.ChucVu from TacGiaBaiBao,GiangVien where MaBB = '" + ma + "' and  GiangVien.MaGV = TacGiaBaiBao.MaGV ";
                DataTable tb = my.DocDL(sql);

                dgv_tacgia.DataSource = tb;
                dgv_tacgia.Columns[0].HeaderText = "Mã giảng viên";
                dgv_tacgia.Columns[1].HeaderText = "Tên giảng viên";
                dgv_tacgia.Columns[2].HeaderText = "Chức vụ";
                dgv_tacgia.Columns[2].Width = 150;




            }
            catch
            {
                MessageBox.Show("Lỗi lấy dữ liệu vào danh sách tác giả", "Thông báo");
            }
        }
        private void uc_baibaoquocte_Load(object sender, EventArgs e)
        {
            LoadDL();
            txt_chucvugv.Text = "Tác giả";
            LoadProductList();
        }

        private void dgv_baibao_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                txt_ma.Text = dgv_baibao.CurrentRow.Cells[0].Value.ToString();
                txt_tenbb.Text = dgv_baibao.CurrentRow.Cells[1].Value.ToString();
                txt_tentapchi.Text = dgv_baibao.CurrentRow.Cells[2].Value.ToString();
                txt_sovatgxb.Text = dgv_baibao.CurrentRow.Cells[3].Value.ToString();
                txt_ISN.Text = dgv_baibao.CurrentRow.Cells[4].Value.ToString();
                txt_namxb.Text = dgv_baibao.CurrentRow.Cells[5].Value.ToString();
                txt_diemct.Text = dgv_baibao.CurrentRow.Cells[6].Value.ToString();
                txt_ghichu.Text = dgv_baibao.CurrentRow.Cells[7].Value.ToString();
                txt_tacgia.Text = dgv_baibao.CurrentRow.Cells[8].Value.ToString();
                
               
                if (e.RowIndex >= 0)
                {
                    object mabb = dgv_baibao.Rows[e.RowIndex].Cells[0].Value;
                    string ma = mabb.ToString();
                    Mabb = ma;
                    loadDLTG(ma);
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(" $ Lỗi cellclick { " + ex.Message + "}", "Thông báo");
            }
        }



        public bool kiemtra()
        {
            if (string.IsNullOrWhiteSpace(txt_ma.Text) || string.IsNullOrWhiteSpace(txt_tenbb.Text)
                || string.IsNullOrWhiteSpace(txt_tentapchi.Text) || string.IsNullOrWhiteSpace(txt_sovatgxb.Text)
                || string.IsNullOrWhiteSpace(txt_ISN.Text) || string.IsNullOrWhiteSpace(txt_namxb.Text )|| string.IsNullOrWhiteSpace(txt_tacgia.Text)
                )
            {
                return false;
            }

            return true;
        }
        public bool kiemtraSo()
        {
            int nam;
            string namm = txt_namxb.Text;
            if ( !int.TryParse(namm,out nam) && namm.Count() > 4)
            {
                return false;
            }

            return true;
        }
        public bool kiemTraMa(string ma)
        {
            try
            {
                string sql = "select * from BaiBao Where MaBB = '" + ma + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }

            }
            catch
            {
                MessageBox.Show("Lỗi không kiểm tra được mã", "Thông báo");

            }
            return true;
        }

        public bool kiemTraMaTG(string ma)
        {
            try
            {
                string sql = "select * from TacGiaBaiBao Where MaBB = '" + ma + "' and MaGV = '" + txt_magv.Text + "' ";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    return false;
                }

            }
            catch
            {
                MessageBox.Show("Lỗi không kiểm tra được mã tác giả", "Thông báo");

            }
            return true;
        }

        private void btn_them_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (kiemtra())
            {
                
                
                if(kiemtraSo())
                {
                    if (kiemTraMa(txt_ma.Text))
                    {

                        //string sql = "insert into BaiBao values('" + txt_ma.Text + "','" + txt_tenbb.Text + "',N'" + txt_tentapchi.Text + "',N'" + txt_sovatgxb.Text + "','" + txt_ISN.Text + "','"+txt_namxb.Text+"','"+txt_diemct.Text+"',N'"+txt_ghichu.Text+"',N'"+LoaiBaiBao+"' )";
                        string sql = "INSERT INTO BaiBao VALUES(@Ma, @TenBB, @TenTapChi, @SoVATGXB, @ISN, @NamXB, @DiemCT, @GhiChu,@TacGia, @LoaiBaiBao)";
                        SqlCommand command = my.SqlCommand(sql);
                        
                            
                            command.Parameters.AddWithValue("@Ma", txt_ma.Text);
                            command.Parameters.AddWithValue("@TenBB", txt_tenbb.Text);
                            command.Parameters.AddWithValue("@TenTapChi", txt_tentapchi.Text);
                            command.Parameters.AddWithValue("@SoVATGXB", txt_sovatgxb.Text);
                            command.Parameters.AddWithValue("@ISN", txt_ISN.Text);
                            command.Parameters.AddWithValue("@NamXB", txt_namxb.Text);
                            command.Parameters.AddWithValue("@DiemCT", txt_diemct.Text);
                            command.Parameters.AddWithValue("@GhiChu", txt_ghichu.Text);
                            command.Parameters.AddWithValue("@TacGia", txt_tacgia.Text);
                            command.Parameters.AddWithValue("@LoaiBaiBao", LoaiBaiBao);
                            //int up = my.Update(sql);
                            int up = command.ExecuteNonQuery();
                        if (up > 0)
                        {
                            MessageBox.Show("Thông tin được thêm thành công", "Thông báo");
                            txt_ma.Clear();
                            txt_tenbb.Clear();
                            txt_tentapchi.Clear();
                            txt_sovatgxb.Clear();
                            txt_ISN.Clear();
                            txt_namxb.Clear();
                            txt_diemct.Clear();
                            txt_ghichu.Clear();
                            txt_tacgia.Clear();
                            LoadDL();
                        }
                        else
                        {
                            MessageBox.Show("Thông tin thêm không thành công", "Thông báo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Đã có mã bài báo này !", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng kiểm tra lại năm xuất bản !", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng điền đủ thông tin !", "Thông báo");
            }
        }

        private void btn_sua_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (kiemtra())
            {

                
                if (kiemtraSo())
                {
                    if (!kiemTraMa(txt_ma.Text))
                    {

                        //string sql = "insert into BaiBao values('" + txt_ma.Text + "','" + txt_tenbb.Text + "',N'" + txt_tentapchi.Text + "',N'" + txt_sovatgxb.Text + "','" + txt_ISN.Text + "','"+txt_namxb.Text+"','"+txt_diemct.Text+"',N'"+txt_ghichu.Text+"',N'"+LoaiBaiBao+"' )";
                        string sql = "update BaiBao set TenBaiBao= @TenBB, TenTapChi=@TenTapChi, SoVaThoiGianXB=@SoVATGXB,ChiSoISN= @ISN,NamXB= @NamXB,DiemCongTrinh= @DiemCT,GhiChu= @GhiChu,TacGia=@TacGia,LoaiBaiBao= @LoaiBaiBao where MaBB=@Ma";
                        SqlCommand command = my.SqlCommand(sql);


                        command.Parameters.AddWithValue("@Ma", txt_ma.Text);
                        command.Parameters.AddWithValue("@TenBB", txt_tenbb.Text);
                        command.Parameters.AddWithValue("@TenTapChi", txt_tentapchi.Text);
                        command.Parameters.AddWithValue("@SoVATGXB", txt_sovatgxb.Text);
                        command.Parameters.AddWithValue("@ISN", txt_ISN.Text);
                        command.Parameters.AddWithValue("@NamXB", txt_namxb.Text);
                        command.Parameters.AddWithValue("@DiemCT", txt_diemct.Text);
                        command.Parameters.AddWithValue("@GhiChu", txt_ghichu.Text);
                        command.Parameters.AddWithValue("@TacGia", txt_tacgia.Text);
                        command.Parameters.AddWithValue("@LoaiBaiBao", LoaiBaiBao);
                        //int up = my.Update(sql);
                        int up = command.ExecuteNonQuery();
                        if (up > 0)
                        {
                            MessageBox.Show("Thông tin được sửa thành công", "Thông báo");
                            txt_ma.Clear();
                            txt_tenbb.Clear();
                            txt_tentapchi.Clear();
                            txt_sovatgxb.Clear();
                            txt_ISN.Clear();
                            txt_namxb.Clear();
                            txt_diemct.Clear();
                            txt_ghichu.Clear();
                            txt_tacgia.Clear();
                            LoadDL();
                        }
                        else
                        {
                            MessageBox.Show("Thông tin sửa không thành công", "Thông báo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("không có bài báo này !", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng kiểm tra lại năm xuất bản !", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn bài báo cần sửa !", "Thông báo");
            }
        }

        private void btn_xoa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                if (kiemtra())
                {
                    if (kiemtraSo())
                    {
                        if (!kiemTraMa(txt_ma.Text))
                        {
                            DialogResult tb = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                            if (tb == DialogResult.OK)
                            {

                                string query = "delete from TacGiaBaiBao where MaBB=@Ma";
                                SqlCommand commandtg = my.SqlCommand(query);
                                commandtg.Parameters.AddWithValue("@Ma", Mabb);

                                int upTG = commandtg.ExecuteNonQuery();

                                if (upTG >= 0)
                                {
                                    string sql = "delete from BaiBao where MaBB=@Ma";
                                    SqlCommand command = my.SqlCommand(sql);
                                    command.Parameters.AddWithValue("@Ma", txt_ma.Text);
                                    int up = command.ExecuteNonQuery();
                                    if (up > 0)
                                    {
                                        MessageBox.Show("Thông tin được xóa thành công", "Thông báo");
                                        txt_ma.Clear();
                                        txt_tenbb.Clear();
                                        txt_tentapchi.Clear();
                                        txt_sovatgxb.Clear();
                                        txt_ISN.Clear();
                                        txt_namxb.Clear();
                                        txt_diemct.Clear();
                                        txt_ghichu.Clear();
                                        txt_tacgia.Clear();
                                        dgv_tacgia.DataSource = null;
                                        LoadDL();
                                    }

                                }
                                else
                                {
                                    MessageBox.Show("Thông tin xóa không thành công", "Thông báo");
                                }
                            }
                            else
                            {

                            }
                        }
                        else
                        {
                            MessageBox.Show("không có bài báo này !", "Thông báo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Vui lòng kiểm tra lại năm xuất bản !", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show("Vui lòng chọn bái báo cần xóa !", "Thông báo");
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi trong quá trình xóa !", "Lỗi");
            }
        }

        private void btn_refesh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadDL();
            txt_ma.Clear();
            txt_tenbb.Clear();
            txt_tentapchi.Clear();
            txt_timkiem.Clear();
            cbo_loai.SelectedIndex = -1;
            dgv_tacgia.DataSource = null;
            txt_sovatgxb.Clear();
            txt_ISN.Clear();
            txt_tengv.Clear();
            txt_namxb.Clear() ;
            txt_diemct.Clear() ;
            txt_ghichu.Clear();
            txt_tacgia.Clear();
        }
        public void ExcelExport()
        {
            try
            {


                string sql = "select MaBB,TenBaiBao,TenTapChi,SoVaThoiGianXB,ChiSoISN,NamXB,DiemCongTrinh,GhiChu from BaiBao where LoaiBaiBao =N'" + LoaiBaiBao + "' ";


                DataTable tb = my.DocDL(sql);

                Excel.Application oExcel = new Excel.Application();
                Excel.Workbook oBook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

                Excel.Worksheet oSheet = (Excel.Worksheet)oBook.Worksheets[1];

                Excel.Range head = oSheet.get_Range("A1", "I1");

                head.MergeCells = true;

                head.Value2 = "DANH SÁCH BÀI BÁO QUỐC TẾ";

                head.Font.Bold = true;

                head.Font.Name = "Times New Roman";

                head.Font.Size = "20";

                head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range cl1 = oSheet.get_Range("A3", "A3");
                cl1.Value = "Mã bài báo";

                Excel.Range cl2 = oSheet.get_Range("B3", "B3");
                cl2.Value = "Tên bài báo";

                Excel.Range cl3 = oSheet.get_Range("C3", "C3");
                cl3.Value = "Tên tạp chí";

                Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                cl4.Value = "Số và thời Gian xuất bản";

                Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                cl5.Value = "Chỉ số ISN";

                Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                cl6.Value = "Năm xuất bản";

                Excel.Range cl7 = oSheet.get_Range("G3", "G3");
                cl7.Value = "Điểm công trình";

                Excel.Range cl8 = oSheet.get_Range("H3", "H3");
                cl8.Value = "Ghi chú";

                Excel.Range cl9 = oSheet.get_Range("I3", "I3");
                cl9.Value = "Họ và Tên tác giả";

                Excel.Range rowHead = oSheet.get_Range("A3", "I3");
                rowHead.Font.Bold = true;
                rowHead.Font.Size = 13;
                rowHead.Font.Name = "Times New Roman";
                rowHead.Borders.LineStyle = Excel.Constants.xlSolid;
                rowHead.Interior.ColorIndex = 6;
                rowHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;




                int line = 4;
                int lines = 4;
                string ma;
                for (int i = 0; i < tb.Rows.Count; i++)
                {

                    for (int j = 0; j < tb.Columns.Count; j++)
                    {
                        oSheet.Cells[i + line, j + 1] = tb.Rows[i][j];
                        oSheet.Cells[i + line, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        oSheet.Cells[i + line, j + 1].Borders.LineStyle = Excel.Constants.xlSolid;
                        oSheet.Cells[i + line, j + 1].Font.Name = "Times New Roman";

                    }


                    ma = tb.Rows[i][0].ToString();
                    string query = "select GiangVien.HoTen from TacGiaBaiBao,GiangVien where GiangVien.MaGV = TacGiaBaiBao.MaGV and MaBB = '" + ma + "' ";
                    DataTable dt = my.DocDL(query);

                    Excel.Range line1 = oSheet.get_Range("I" + (lines).ToString(), "I" + (lines).ToString());

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {

                        string cel = dt.Rows[row]["HoTen"].ToString() + "\n";
                        line1.Value += cel;


                    }
                    line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line1.Borders.LineStyle = Excel.Constants.xlSolid;
                    line1.Font.Name = "Times New Roman";
                    lines++;
                }

                oSheet.Name = "BBQT";
                oExcel.Columns.AutoFit();

                oBook.Activate();

                SaveFileDialog saveFile = new SaveFileDialog();
                if (saveFile.ShowDialog() == DialogResult.OK)
                {

                    saveFile.Filter = "Text files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    oBook.SaveAs(saveFile.FileName.ToLower());
                    MessageBox.Show("Xuất danh sách thành công");

                }

                oExcel.Quit();

            }
            catch
            {
                MessageBox.Show("Xuất danh sách không thành công");
            }
        }
        private void btn_export_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            
        }

        private void btn_timkiem_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(txt_timkiem.Text))
            {
                if(cbo_loai.SelectedIndex != -1)
                {
                    int i = cbo_loai.SelectedIndex;
                    string sql;

                    switch (i)
                    {
                        case 0:
                            try
                            {
                                sql = "select MaBB,TenBaiBao,TenTapChi,SoVaThoiGianXB,ChiSoISN,NamXB,DiemCongTrinh,GhiChu,TacGia from BaiBao where LoaiBaiBao =N'" + LoaiBaiBao + "' and MaBB like '%" + txt_timkiem.Text + "%' ";
                                DataTable tb = my.DocDL(sql);

                                dgv_baibao.DataSource = tb;
                                dgv_baibao.Columns[0].HeaderText = "Mã bài báo";
                                dgv_baibao.Columns[1].HeaderText = "Tên bài báo";
                                dgv_baibao.Columns[1].Width = 250;
                                dgv_baibao.Columns[2].HeaderText = "Tên tạp chí";
                                dgv_baibao.Columns[2].Width = 250;
                                dgv_baibao.Columns[3].HeaderText = "Số vào Thời gian xuất bản";
                                dgv_baibao.Columns[3].Width = 150;
                                dgv_baibao.Columns[4].HeaderText = "Chỉ số ISN";
                                dgv_baibao.Columns[5].HeaderText = "Năm xuất bản";
                                dgv_baibao.Columns[6].HeaderText = "Điểm công trình";
                                dgv_baibao.Columns[7].HeaderText = "Ghi chú";
                                dgv_baibao.Columns[8].HeaderText = "Tác giả";
                            }
                            catch
                            {
                                MessageBox.Show("Lỗi tìm kiếm theo mã bài báo ", "Thông báo");
                            }
                            
                            break;
                        case 1:
                            try
                            {
                                sql = "select MaBB,TenBaiBao,TenTapChi,SoVaThoiGianXB,ChiSoISN,NamXB,DiemCongTrinh,GhiChu,TacGia from BaiBao where LoaiBaiBao =N'" + LoaiBaiBao + "' and TenBaiBao like '%'+@MaBB+'%' ";
                                SqlCommand command = my.SqlCommand(sql);
                                command.Parameters.AddWithValue("@MaBB", txt_timkiem.Text);

                                SqlDataAdapter adapter = new SqlDataAdapter(command);
                                DataTable dataTable = new DataTable();
                                adapter.Fill(dataTable);

                                dgv_baibao.DataSource = dataTable;

                                dgv_baibao.Columns[0].HeaderText = "Mã bài báo";
                                dgv_baibao.Columns[1].HeaderText = "Tên bài báo";
                                dgv_baibao.Columns[1].Width = 250;
                                dgv_baibao.Columns[2].HeaderText = "Tên tạp chí";
                                dgv_baibao.Columns[2].Width = 250;
                                dgv_baibao.Columns[3].HeaderText = "Số vào Thời gian xuất bản";
                                dgv_baibao.Columns[3].Width = 150;
                                dgv_baibao.Columns[4].HeaderText = "Chỉ số ISN";
                                dgv_baibao.Columns[5].HeaderText = "Năm xuất bản";
                                dgv_baibao.Columns[6].HeaderText = "Điểm công trình";
                                dgv_baibao.Columns[7].HeaderText = "Ghi chú";
                                dgv_baibao.Columns[8].HeaderText = "Tác giả";
                            }
                            catch
                            {
                                MessageBox.Show("Lỗi tìm kiếm theo tên bài báo ", "Thông báo");
                            }

                            break;
                    }



                }
                else
                {
                    MessageBox.Show("Vui lòng chọn nội dung tìm kiếm ", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng nhập thông tin tìm kiếm ", "Thông báo");
            }
        }

        private void dgv_tacgia_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_magv.Text = dgv_tacgia.CurrentRow.Cells[0].Value.ToString();
                txt_tengv.Text = dgv_tacgia.CurrentRow.Cells[1].Value.ToString();
                //txt_chucvugv.Text = dgv_tacgia.CurrentRow.Cells[2].Value.ToString();



            }
            catch (Exception ex)
            {
                MessageBox.Show(" $ Lỗi cellclick { " + ex.Message + "}", "Thông báo");
            }
        }
        private void LoadProductList()
        {
            try
            {

                productList = new List<string>();
                string query = "SELECT MaGV FROM GiangVien";
                DataTable tb = my.DocDL(query);
                if (tb.Rows.Count > 0)
                {
                    for (int i = 0; i < tb.Rows.Count; i++)
                    {
                        string ma = tb.Rows[i][0].ToString();
                        productList.Add(ma);
                    }
                }
            }
            catch
            {
                MessageBox.Show($"Lỗi thực hiện tạo danh sách ", "Lỗi");
            }

        }
        private void ShowSuggestions(List<string> suggestions)
        {
            list_gv.Items.Clear();
            list_gv.Items.AddRange(suggestions.ToArray());

            list_gv.Visible = suggestions.Any();
        }
        private void txt_magv_TextChanged(object sender, EventArgs e)
        {
            if (dgv_tacgia.Rows.Count >= 0)
            {
                string searchTerm = txt_magv.Text.ToLower();
                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
                    List<string> filteredProducts = productList
                   .Where(product => product.ToLower().Contains(searchTerm))
                   .ToList();

                    ShowSuggestions(filteredProducts);
                }
                else
                {
                    list_gv.Visible = false;
                }


            }
            else
            {
                MessageBox.Show($"Vui lòng chọn bài báo ", "Thông báo");
            }
        }

        private void list_gv_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (list_gv.SelectedItem != null)
            {
                string selectedProduct = list_gv.SelectedItem.ToString();
                if (!string.IsNullOrWhiteSpace(selectedProduct))
                {
                    txt_magv.Text = selectedProduct;
                    list_gv.Visible = false;
                    string sql = "select HoTen from GiangVien where MaGV = '" + selectedProduct + "' ";
                    DataTable tb = my.DocDL(sql);
                    if (tb.Rows.Count > 0)
                    {
                        string hoten = tb.Rows[0][0].ToString();
                        txt_tengv.Text = hoten;
                    }
                }

            }
        }

        private void btn_joingv_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_magv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text) || string.IsNullOrWhiteSpace(txt_chucvugv.Text))
            {
                MessageBox.Show(" $ Vui lòng nhập đầy đủ thông tin ", "Thông báo");
            }
            else
            {

                if (dgv_tacgia.DataSource == null)
                {
                    MessageBox.Show("Vui lòng chọn bài báo muốn thêm tác giả !", "Thông báo");
                }
                else
                {

                    string ma = Mabb;
                    if (kiemTraMaTG(txt_ma.Text))
                    {

                        string sql = "insert into TacGiaBaiBao values('" + ma + "','" + txt_magv.Text + "',N'" + txt_chucvugv.Text + "')";
                        int up = my.Update(sql);
                        if (up > 0)
                        {
                            MessageBox.Show("Thông tin tác giả được thêm thành công", "Thông báo");
                            txt_magv.Clear();
                            txt_tengv.Clear();
                            


                            loadDLTG(ma);
                        }
                        else
                        {
                            MessageBox.Show("Thông tin tác giả thêm không thành công", "Thông báo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Đã có tác giả này !", "Thông báo");
                    }
                }

            }
        }

        private void btn_cancelgv_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_magv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text) || string.IsNullOrWhiteSpace(txt_chucvugv.Text))
            {
                MessageBox.Show(" $ Vui lòng nhập đầy đủ thông tin ", "Thông báo");
            }
            else
            {

                if (dgv_tacgia.DataSource == null)
                {
                    MessageBox.Show("Vui lòng chọn bài báo muốn xóa tác giả !", "Thông báo");
                }
                else
                {

                    string ma = Mabb;
                    if (!kiemTraMaTG(ma))
                    {

                        string sql = "delete from TacGiaBaiBao where MaBB='" + ma + "'and MaGV='" + txt_magv.Text + "' ";
                        int up = my.Update(sql);
                        if (up > 0)
                        {
                            MessageBox.Show("Xóa thông tin thành công", "Thông báo");
                            txt_magv.Clear();
                            txt_tengv.Clear();
                            


                            loadDLTG(ma);
                        }
                        else
                        {
                            MessageBox.Show("Xóa thông tin không thành công", "Thông báo");
                        }
                    }
                    else
                    {
                        MessageBox.Show("không phải tác giả của tạp chí !", "Thông báo");
                    }
                }

            }
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ExcelExport();
        }
        public void ExcelExport1BB()
        {
            try
            {
                if(kiemtra())
                {
                    string sql = "select MaBB,TenBaiBao,TenTapChi,SoVaThoiGianXB,ChiSoISN,NamXB,DiemCongTrinh,GhiChu from BaiBao where LoaiBaiBao =N'" + LoaiBaiBao + "' and MaBB='" + txt_ma.Text + "' ";


                    DataTable tb = my.DocDL(sql);

                    Excel.Application oExcel = new Excel.Application();
                    Excel.Workbook oBook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

                    Excel.Worksheet oSheet = (Excel.Worksheet)oBook.Worksheets[1];

                    Excel.Range head = oSheet.get_Range("A1", "I1");

                    head.MergeCells = true;

                    head.Value2 = "DANH SÁCH BÀI BÁO QUỐC TẾ";

                    head.Font.Bold = true;

                    head.Font.Name = "Times New Roman";

                    head.Font.Size = "20";

                    head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    Excel.Range cl1 = oSheet.get_Range("A3", "A3");
                    cl1.Value = "Mã bài báo";

                    Excel.Range cl2 = oSheet.get_Range("B3", "B3");
                    cl2.Value = "Tên bài báo";

                    Excel.Range cl3 = oSheet.get_Range("C3", "C3");
                    cl3.Value = "Tên tạp chí";

                    Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                    cl4.Value = "Số và thời Gian xuất bản";

                    Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                    cl5.Value = "Chỉ số ISN";

                    Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                    cl6.Value = "Năm xuất bản";

                    Excel.Range cl7 = oSheet.get_Range("G3", "G3");
                    cl7.Value = "Điểm công trình";

                    Excel.Range cl8 = oSheet.get_Range("H3", "H3");
                    cl8.Value = "Ghi chú";

                    Excel.Range cl9 = oSheet.get_Range("I3", "I3");
                    cl9.Value = "Họ và Tên tác giả";

                    Excel.Range rowHead = oSheet.get_Range("A3", "I3");
                    rowHead.Font.Bold = true;
                    rowHead.Font.Size = 13;
                    rowHead.Font.Name = "Times New Roman";
                    rowHead.Borders.LineStyle = Excel.Constants.xlSolid;
                    rowHead.Interior.ColorIndex = 6;
                    rowHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;




                    int line = 4;
                    int lines = 4;
                    string ma;
                    for (int i = 0; i < tb.Rows.Count; i++)
                    {

                        for (int j = 0; j < tb.Columns.Count; j++)
                        {
                            oSheet.Cells[i + line, j + 1] = tb.Rows[i][j];
                            oSheet.Cells[i + line, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            oSheet.Cells[i + line, j + 1].Borders.LineStyle = Excel.Constants.xlSolid;
                            oSheet.Cells[i + line, j + 1].Font.Name = "Times New Roman";

                        }


                        ma = tb.Rows[i][0].ToString();
                        string query = "select GiangVien.HoTen from TacGiaBaiBao,GiangVien where GiangVien.MaGV = TacGiaBaiBao.MaGV and MaBB = '" + ma + "' ";
                        DataTable dt = my.DocDL(query);

                        Excel.Range line1 = oSheet.get_Range("I" + (lines).ToString(), "I" + (lines).ToString());

                        for (int row = 0; row < dt.Rows.Count; row++)
                        {

                            string cel = dt.Rows[row]["HoTen"].ToString() + "\n";
                            line1.Value += cel;


                        }
                        line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line1.Borders.LineStyle = Excel.Constants.xlSolid;
                        line1.Font.Name = "Times New Roman";
                        lines++;
                    }

                    oSheet.Name = "BBQT";
                    oExcel.Columns.AutoFit();

                    oBook.Activate();

                    SaveFileDialog saveFile = new SaveFileDialog();
                    if (saveFile.ShowDialog() == DialogResult.OK)
                    {

                        saveFile.Filter = "Text files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                        oBook.SaveAs(saveFile.FileName.ToLower());
                        MessageBox.Show("Xuất danh sách thành công");

                    }

                    oExcel.Quit();
                }else
                {
                    MessageBox.Show("Vui lòng chọn bài báo muốn export dữ liệu","Thông báo");
                }
            


            }
            catch
            {
                MessageBox.Show("Xuất danh sách không thành công");
            }
        }
        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ExcelExport1BB();
        }
    }
}
