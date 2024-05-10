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
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;
using System.Data.SqlClient;

namespace QL_NCKH
{
    public partial class uc_tapchiNCKH : DevExpress.XtraEditors.XtraUserControl
    {
        MyClass my = new MyClass();
        private List<string> productList;
        private string matc;
        public uc_tapchiNCKH()
        {
            InitializeComponent();
        }
        public string Matc
        {
            get { return this.matc; }
            set { this.matc = value; }
        }
        public void LoadDL()
        {
            try
            {
                string sql = "select * from TapChi";
                DataTable tb = my.DocDL(sql);
                if (tb.Rows.Count > 0)
                {
                    dgv_tapchi.DataSource = tb;
                    dgv_tapchi.Columns[0].HeaderText = "Mã tạp chí";
                    dgv_tapchi.Columns[1].HeaderText = "Tên tạp chí";
                    dgv_tapchi.Columns[1].Width = 250;
                    dgv_tapchi.Columns[2].HeaderText = "Ngày công bố";
                    dgv_tapchi.Columns[3].HeaderText = "Giấy phép xuất bản";
                    dgv_tapchi.Columns[3].Width = 150;
                    dgv_tapchi.Columns[4].HeaderText = "Link công bố trực tiếp";
                    dgv_tapchi.Columns[4].Width = 250;


                }
            }
            catch
            {
                MessageBox.Show("Lỗi lấy dữ liệu vào danh sách", "Thông báo");
            }
        }
        public void loadDLTG(string ma)
        {
            try
            {

                string sql = "select GiangVien.MaGV,GiangVien.HoTen,CTTapChi.ChucVu from CTTapChi,GiangVien where MaTC = '" + ma + "' and  GiangVien.MaGV = CTTapChi.MaGV ";
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

        private void uc_tapchiNCKH_Load(object sender, EventArgs e)
        {
            LoadDL();
            LoadProductList();
        }

        private void dgv_tapchi_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                txt_ma.Text = dgv_tapchi.CurrentRow.Cells[0].Value.ToString();
                txt_tieude.Text = dgv_tapchi.CurrentRow.Cells[1].Value.ToString();
                dtp_ngay.Text = dgv_tapchi.CurrentRow.Cells[2].Value.ToString();
                txt_giayphep.Text = dgv_tapchi.CurrentRow.Cells[3].Value.ToString();
                txt_link.Text = dgv_tapchi.CurrentRow.Cells[4].Value.ToString();
                if (e.RowIndex >= 0)
                {
                    object matc = dgv_tapchi.Rows[e.RowIndex].Cells[0].Value;
                    string ma = matc.ToString();
                    Matc = ma;
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
            if (string.IsNullOrWhiteSpace(txt_ma.Text) || string.IsNullOrWhiteSpace(txt_tieude.Text)
                || string.IsNullOrWhiteSpace(txt_link.Text) || string.IsNullOrWhiteSpace(txt_giayphep.Text))
            {
                return false;
            }

            return true;
        }
        public bool kiemTraMa(string ma)
        {
            try
            {
                string sql = "select * from TapChi Where MaTC = '" + ma + "' ";
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
                string sql = "select * from CTTapChi Where MaTC = '" + ma + "' and MaGV = '" + txt_magv.Text + "' ";
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
                if (kiemTraMa(txt_ma.Text))
                {
                    string ngaycb = dtp_ngay.Value.ToString("yyyy/MM/dd");
                    string sql = "insert into TapChi values('" + txt_ma.Text + "',N'" + txt_tieude.Text + "','" + ngaycb + "',N'" + txt_giayphep.Text + "','" + txt_link.Text + "')";
                    int up = my.Update(sql);
                    if (up > 0)
                    {
                        MessageBox.Show("Thông tin được thêm thành công", "Thông báo");
                        txt_ma.Clear();
                        txt_tieude.Clear();
                        txt_link.Clear();
                        txt_giayphep.Clear();

                        LoadDL();
                    }
                    else
                    {
                        MessageBox.Show("Thông tin thêm không thành công", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show("Đã có mã tạp chí này !", "Thông báo");
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
                if (!kiemTraMa(txt_ma.Text))
                {
                    string ngaycb = dtp_ngay.Value.ToString("yyyy/MM/dd");
                    string sql = "update TapChi set TieuDeCongBo=N'" + txt_tieude.Text + "',NgayCongBo='" + ngaycb + "',GiayPhep=N'" + txt_giayphep.Text + "',LinkCongBoTrucTiep=N'" + txt_link.Text + "' where MaTC = '" + txt_ma.Text + "' ";
                    int up = my.Update(sql);
                    if (up > 0)
                    {
                        MessageBox.Show("Thông tin được sửa thành công", "Thông báo");
                        txt_ma.Clear();
                        txt_tieude.Clear();
                        txt_link.Clear();
                        txt_giayphep.Clear();

                        LoadDL();
                    }
                    else
                    {
                        MessageBox.Show("Thông tin sửa không thành công", "Thông báo");
                    }
                }
                else
                {
                    MessageBox.Show("Không có mã tạp chí này !", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn tạp chí muốn sửa !", "Thông báo");
            }
        }

        private void btn_xoa_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (kiemtra())
            {
                if (!kiemTraMa(txt_ma.Text))
                {

                    DialogResult tb = MessageBox.Show("Xin lưu ý rằng hành động này sẽ xóa một số dữ liệu quan trọng. Bạn có chắc chắn muốn tiếp tục?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

                    if (tb == DialogResult.OK)
                    {
                        string query = "delete from CTTapChi where MaTC = '" + Matc + "' ";
                        int up1 = my.Update(query);
                        if (up1 >= 0)
                        {
                            string sql = "delete from TapChi where MaTC = '" + txt_ma.Text + "' ";
                            int up = my.Update(sql);
                            if (up > 0)
                            {
                                MessageBox.Show("Thông tin được xóa thành công", "Thông báo");
                                txt_ma.Clear();
                                txt_tieude.Clear();
                                txt_link.Clear();
                                txt_giayphep.Clear();
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
                    MessageBox.Show("không có mã tạp chí này !", "Thông báo");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn tạp chí muốn xóa !", "Thông báo");
            }
        }

        private void btn_timkiem_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
        }

        private void btn_refesh_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            LoadDL();
            txt_ma.Clear();
            txt_link.Clear();
            txt_tieude.Clear();
            txt_timkiem.Clear();
            cbo_loai.SelectedIndex = -1;
            dgv_tacgia.DataSource = null;
            txt_magv.Clear();
            txt_giayphep.Clear();
            txt_tengv.Clear();
            cbo_chucvugv.SelectedIndex = -1;
        }
        public void ExcelExport()
        {
            try
            {


                string sql = @"select * from  TapChi";
                

                DataTable tb = my.DocDL(sql);

                Excel.Application oExcel = new Excel.Application();
                Excel.Workbook oBook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

                Excel.Worksheet oSheet = (Excel.Worksheet)oBook.Worksheets[1];

                Excel.Range head = oSheet.get_Range("A1", "F1");

                head.MergeCells = true;

                head.Value2 = "Tạp chí NCKH";

                head.Font.Bold = true;

                head.Font.Name = "Times New Roman";

                head.Font.Size = "20";

                head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Excel.Range cl1 = oSheet.get_Range("A3", "A3");
                cl1.Value = "Mã tạp chí";

                Excel.Range cl2 = oSheet.get_Range("B3", "B3");
                cl2.Value = "Tiêu đề công bố";
                Excel.Range cl3 = oSheet.get_Range("C3", "C3");
                cl3.Value = "Ngày công bố";

                Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                cl4.Value = "Giấy phép xuất bản";

                Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                cl5.Value = "Link công bố trực tiếp";

                Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                cl6.Value = "Tác giả";

                //Excel.Range cl7 = oSheet.get_Range("G3", "G3");
                //cl7.Value = "Chức vụ";


                Excel.Range rowHead = oSheet.get_Range("A3", "F3");
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
                    string query = "select GiangVien.HoTen,CTTapChi.ChucVu from CTTapChi,GiangVien where GiangVien.MaGV = CTTapChi.MaGV and MaTC = '" + ma + "' ";
                    DataTable dt = my.DocDL(query);

                    Excel.Range line1 = oSheet.get_Range("F" + (lines).ToString(), "F" + (lines).ToString());
                    
                    for (int row = 0; row < dt.Rows.Count; row++)
                    {

                        string cel = dt.Rows[row]["HoTen"].ToString() + "-" + dt.Rows[row]["ChucVu"].ToString() + "\n";
                        line1.Value += cel;
                        

                    }
                    line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    line1.Borders.LineStyle = Excel.Constants.xlSolid;
                    line1.Font.Name = "Times New Roman"; 
                    lines++;
                }

                oSheet.Name = "TCNCKH";
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


        private void OpenWebPageInPreferredBrowser(string url)
        {
            string[] browserPaths = { "chrome.exe", "CocCoc.exe", "firefox.exe", "iexplore.exe" };

            foreach (var browserPath in browserPaths)
            {
                if (IsBrowserInstalled(browserPath))
                {
                    try
                    {
                        Process.Start(browserPath, url);
                        return;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Lỗi mở trang wed: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            MessageBox.Show("Không có trình duyệt nào trong máy .", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private bool IsBrowserInstalled(string browserPath)
        {
            return Process.GetProcessesByName(Path.GetFileNameWithoutExtension(browserPath)).Any();
        }
        private void btn_open_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string url = txt_link.Text;

            if (!string.IsNullOrEmpty(url))
            {
                OpenWebPageInPreferredBrowser(url);
            }
        }

        private void dvg_tacgia_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txt_magv.Text = dgv_tacgia.CurrentRow.Cells[0].Value.ToString();
                txt_tengv.Text = dgv_tacgia.CurrentRow.Cells[1].Value.ToString();
                cbo_chucvugv.Text = dgv_tacgia.CurrentRow.Cells[2].Value.ToString();



            }
            catch (Exception ex)
            {
                MessageBox.Show(" $ Lỗi cellclick { " + ex.Message + "}", "Thông báo");
            }
        }

        private void btn_joingv_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_magv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text) || string.IsNullOrWhiteSpace(cbo_chucvugv.Text))
            {
                MessageBox.Show(" $ Vui lòng nhập đầy đủ thông tin ", "Thông báo");
            }
            else
            {

                if (dgv_tacgia.DataSource == null)
                {
                    MessageBox.Show("Vui lòng chọn tạp chí muốn thêm tác giả !", "Thông báo");
                }
                else
                {

                    string ma = Matc;
                    if (kiemTraMaTG(txt_ma.Text))
                    {

                        string sql = "insert into CTTapChi values('" + ma + "','" + txt_magv.Text + "',N'" + cbo_chucvugv.Text + "')";
                        int up = my.Update(sql);
                        if (up > 0)
                        {
                            MessageBox.Show("Thông tin tác giả được thêm thành công", "Thông báo");
                            txt_magv.Clear();
                            txt_tengv.Clear();
                            cbo_chucvugv.SelectedIndex = -1;


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
                MessageBox.Show($"Vui lòng chọn tạp chí ", "Thông báo");
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

        private void dgv_tapchi_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btn_cancelgv_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_magv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text) || string.IsNullOrWhiteSpace(cbo_chucvugv.Text))
            {
                MessageBox.Show(" $ Vui lòng nhập đầy đủ thông tin ", "Thông báo");
            }
            else
            {

                if (dgv_tacgia.DataSource == null)
                {
                    MessageBox.Show("Vui lòng chọn tạp chí muốn xóa tác giả !", "Thông báo");
                }
                else
                {

                    string ma = Matc;
                    if (!kiemTraMaTG(ma))
                    {

                        string sql = "delete from CTTapChi where MaTC='" + ma + "'and MaGV='" + txt_magv.Text + "' ";
                        int up = my.Update(sql); 
                        if (up > 0)
                        {
                            MessageBox.Show("Xóa thông tin thành công", "Thông báo");
                            txt_magv.Clear();
                            txt_tengv.Clear();
                            cbo_chucvugv.SelectedIndex = -1;


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

        private void btn_suagv_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_magv.Text) || string.IsNullOrWhiteSpace(txt_tengv.Text) || string.IsNullOrWhiteSpace(cbo_chucvugv.Text))
            {
                MessageBox.Show(" $ Vui lòng nhập đầy đủ thông tin ", "Thông báo");
            }
            else
            {

                if (dgv_tacgia.DataSource == null)
                {
                    MessageBox.Show("Vui lòng chọn tạp chí muốn sửa tác giả !", "Thông báo");
                }
                else
                {

                    string ma = Matc;
                    if (!kiemTraMaTG(txt_ma.Text))
                    {

                        string sql = "update CTTapChi set ChucVu = N'"+cbo_chucvugv.Text+"' where MaTC='" + ma + "'and MaGV='" + txt_magv.Text + "' ";
                        int up = my.Update(sql);
                        if (up > 0)
                        {
                            MessageBox.Show("Sửa thông tin thành công", "Thông báo");
                            txt_magv.Clear();
                            txt_tengv.Clear();
                            cbo_chucvugv.SelectedIndex = -1;


                            loadDLTG(ma);
                        }
                        else
                        {
                            MessageBox.Show("Sửa thông tin không thành công", "Thông báo");
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
        public void ExcelExport1TC()
        {
            try
            {
                if(kiemtra())
                {
                    string sql = @"select * from  TapChi where MaTC = '" + txt_ma.Text + "' ";


                    DataTable tb = my.DocDL(sql);

                    Excel.Application oExcel = new Excel.Application();
                    Excel.Workbook oBook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

                    Excel.Worksheet oSheet = (Excel.Worksheet)oBook.Worksheets[1];

                    Excel.Range head = oSheet.get_Range("A1", "F1");

                    head.MergeCells = true;

                    head.Value2 = "Thông Tin Tạp chí NCKH";

                    head.Font.Bold = true;

                    head.Font.Name = "Times New Roman";

                    head.Font.Size = "20";

                    head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    Excel.Range cl1 = oSheet.get_Range("A3", "A3");
                    cl1.Value = "Mã tạp chí";

                    Excel.Range cl2 = oSheet.get_Range("B3", "B3");
                    cl2.Value = "Tiêu đề công bố";
                    Excel.Range cl3 = oSheet.get_Range("C3", "C3");
                    cl3.Value = "Ngày công bố";

                    Excel.Range cl4 = oSheet.get_Range("D3", "D3");
                    cl4.Value = "Giấy phép xuất bản";

                    Excel.Range cl5 = oSheet.get_Range("E3", "E3");
                    cl5.Value = "Link công bố trực tiếp";

                    Excel.Range cl6 = oSheet.get_Range("F3", "F3");
                    cl6.Value = "Tác giả";

                    //Excel.Range cl7 = oSheet.get_Range("G3", "G3");
                    //cl7.Value = "Chức vụ";


                    Excel.Range rowHead = oSheet.get_Range("A3", "F3");
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
                        string query = "select GiangVien.HoTen,CTTapChi.ChucVu from CTTapChi,GiangVien where GiangVien.MaGV = CTTapChi.MaGV and MaTC = '" + ma + "' ";
                        DataTable dt = my.DocDL(query);

                        Excel.Range line1 = oSheet.get_Range("F" + (lines).ToString(), "F" + (lines).ToString());

                        for (int row = 0; row < dt.Rows.Count; row++)
                        {

                            string cel = dt.Rows[row]["HoTen"].ToString() + "-" + dt.Rows[row]["ChucVu"].ToString() + "\n";
                            line1.Value += cel;


                        }
                        line1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        line1.Borders.LineStyle = Excel.Constants.xlSolid;
                        line1.Font.Name = "Times New Roman";
                        lines++;
                    }

                    oSheet.Name = "TCNCKH";
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
                else
                {
                    MessageBox.Show("Vui lòng chọn tạp chí cần export dữ liệu","Thông báo");
                }

                

            }
            catch
            {
                MessageBox.Show("Xuất danh sách không thành công");
            }
        }
        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ExcelExport1TC();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (!string.IsNullOrWhiteSpace(txt_timkiem.Text))
            {
                int i = cbo_loai.SelectedIndex;
                string sql;

                switch (i)
                {
                    case 0:
                        sql = "select * from TapChi where MaTC like '%" + txt_timkiem.Text + "%' ";
                        DataTable tb = my.DocDL(sql);
                        if (tb.Rows.Count > 0)
                        {
                            dgv_tapchi.DataSource = tb;
                            dgv_tapchi.Columns[0].HeaderText = "Mã tạp chí";
                            dgv_tapchi.Columns[1].HeaderText = "Tên tạp chí";
                            dgv_tapchi.Columns[1].Width = 250;
                            dgv_tapchi.Columns[2].HeaderText = "Ngày công bố";
                            dgv_tapchi.Columns[3].HeaderText = "Link công bố trực tiếp";
                            dgv_tapchi.Columns[3].Width = 250;


                        }
                        break;
                    case 1:
                        sql = "select * from TapChi where TieuDeCongBo like N'%" + txt_timkiem.Text + "%' ";
                        DataTable tb1 = my.DocDL(sql);
                        if (tb1.Rows.Count > 0)
                        {
                            dgv_tapchi.DataSource = tb1;
                            dgv_tapchi.Columns[0].HeaderText = "Mã tạp chí";
                            dgv_tapchi.Columns[1].HeaderText = "Tên tạp chí";
                            dgv_tapchi.Columns[1].Width = 250;
                            dgv_tapchi.Columns[2].HeaderText = "Ngày công bố";
                            dgv_tapchi.Columns[3].HeaderText = "Link công bố trực tiếp";
                            dgv_tapchi.Columns[3].Width = 250;


                        }
                        break;


                }
            }
            else
            {
                MessageBox.Show("Vui lòng nhập thông tin tìm kiếm ", "Thông báo");
            }
        }
    }
}
