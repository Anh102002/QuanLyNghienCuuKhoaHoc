using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Excel = Microsoft.Office.Interop.Excel;
namespace QL_NCKH
{
    public partial class frm_thongketheonam : DevExpress.XtraEditors.XtraForm
    {
        MyClass my = new MyClass();
        public frm_thongketheonam()
        {
            InitializeComponent();
        }
        public DataTable LayDuLieuBaoCao()
        {
            int nam = Convert.ToInt32( cbo_nam.Text);
            string query = " SELECT * FROM dbo.ReportDTYear(N'"+cbo_khoa.Text+"', '"+nam+"') ";
            DataTable dataTable = my.DocDL(query);

            return dataTable;
        }
        public void excelCTHT()
        {
            try
            {

                DataTable dataTable = LayDuLieuBaoCao();


                Excel.Application oExcel = new Excel.Application();
                Excel.Workbook workbook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet oSheet = (Excel.Worksheet)workbook.Worksheets[1];




                Excel.Range head = oSheet.get_Range("A1", "G1");

                head.MergeCells = true;

                head.Value2 = "BÁO CÁO ĐỀ TÀI NGHIÊN CỨU CỦA KHOA THEO NĂM";

                head.Font.Bold = true;

                head.Font.Name = "Times New Roman";

                head.Font.Size = "20";
                head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ////

                Excel.Range head1 = oSheet.get_Range("A2", "G2");

                head1.MergeCells = true;
                string ngay = DateTime.Now.ToString("dd-MM-yyyy");
                head1.Value2 = "Ngày lập báo cáo : "+ngay;

                head1.Font.Bold = true;

                head1.Font.Name = "Times New Roman";

                head1.Font.Size = "13";

                //head1.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                /////

                Excel.Range head2 = oSheet.get_Range("A3", "G3");

                head2.MergeCells = true;
                
                head2.Value2 = "Khoa : " + cbo_khoa.Text;

                head2.Font.Bold = true;

                head2.Font.Name = "Times New Roman";

                head2.Font.Size = "13";

                //head2.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                /////


                Excel.Range head3 = oSheet.get_Range("A4", "G4");

                head3.MergeCells = true;

                head3.Value2 = "Năm : " + cbo_nam.Text;

                head3.Font.Bold = true;

                head3.Font.Name = "Times New Roman";

                head3.Font.Size = "13";

                //head3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                /////

                Excel.Range cl1 = oSheet.get_Range("A6", "A6");
                cl1.Value = "Mã đề tài";

                Excel.Range cl2 = oSheet.get_Range("B6", "B6");
                cl2.Value = "Tên đề tài";

                Excel.Range cl3 = oSheet.get_Range("C6", "C6");
                cl3.Value = "Lĩnh vực";

                Excel.Range cl4 = oSheet.get_Range("D6", "D6");
                cl4.Value = "Ngày bắt đầu";

                Excel.Range cl5 = oSheet.get_Range("E6", "E6");
                cl5.Value = "Ngày kết thúc";

                Excel.Range cl10 = oSheet.get_Range("F6", "F6");
                cl10.Value = "Tiến độ";

                Excel.Range cl6 = oSheet.get_Range("G6", "G6");
                cl6.Value = "Cấp đề tài";

                //Excel.Range cl7 = oSheet.get_Range("H3", "H3");
                //cl7.Value = "Ban tổ chức hội thảo";

                //Excel.Range cl8 = oSheet.get_Range("I3", "I3");
                //cl8.Value = "";

                //Excel.Range cl9 = oSheet.get_Range("J3", "J3");
                //cl9.Value = "Giảng viên hướng dẫn";





                Excel.Range rowHead = oSheet.get_Range("A6", "G6");
                rowHead.Font.Bold = true;
                rowHead.Font.Name = "Times New Roman";
                rowHead.Font.Size = 13;
                rowHead.Borders.LineStyle = Excel.Constants.xlSolid;
                rowHead.Interior.ColorIndex = 6;
                rowHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Sau đó, thêm dữ liệu từ DataTable
                int line = 7;
               
                
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {

                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        oSheet.Cells[i + line, j + 1] = dataTable.Rows[i][j];
                        oSheet.Cells[i + line, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        oSheet.Cells[i + line, j + 1].Borders.LineStyle = Excel.Constants.xlSolid;
                        oSheet.Cells[i + line, j + 1].Font.Name = "Times New Roman";

                    }

                }
                /////Tổng đề tài thamg gia
                int dong = line + dataTable.Rows.Count;
                Excel.Range headSum = oSheet.get_Range("A"+dong, "G"+dong);

                headSum.MergeCells = true;

                headSum.Value2 = "Tổng đề tài : " + dataTable.Rows.Count;

                headSum.Font.Bold = true;

                headSum.Font.Name = "Times New Roman";

                headSum.Font.Size = "13";
                headSum.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                headSum.Borders.LineStyle = Excel.Constants.xlSolid;
                //////

                oSheet.Name = "BCDTK";
                oExcel.Columns.AutoFit();

                workbook.Activate();
                SaveFileDialog saveFile = new SaveFileDialog();
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    saveFile.Filter = "Text Files|*.xlxs|All Files|*.*";
                    workbook.SaveAs(saveFile.FileName.ToLower());
                    MessageBox.Show("Xuất danh sách thành công", "Thông báo");
                }

                oExcel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi xuất báo cáo: {ex.Message}");
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(cbo_nam.Text) || string.IsNullOrWhiteSpace(cbo_khoa.Text))
            {
                MessageBox.Show("Vui lòng chọn thông tin xuất báo cáo", "Thông báo");
            }
            else
            {
                excelCTHT();
            }
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            this.Close();
        }
        public DataTable LayDuLieuBaoCaoGV()
        {
            int nam = Convert.ToInt32(cbo_namgv.Text);
            string query = " select * from ReportGVYear('"+txt_magv.Text+"','"+nam+"') ";
            DataTable dataTable = my.DocDL(query);

            return dataTable;
        }
        public void excelGVDT()
        {
            try
            {

                DataTable dataTable = LayDuLieuBaoCaoGV();


                Excel.Application oExcel = new Excel.Application();
                Excel.Workbook workbook = oExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet oSheet = (Excel.Worksheet)workbook.Worksheets[1];




                Excel.Range head = oSheet.get_Range("A1", "H1");

                head.MergeCells = true;

                head.Value2 = "BÁO CÁO ĐỀ TÀI NGHIÊN CỨU CỦA GIẢNG VIÊN THEO NĂM";

                head.Font.Bold = true;

                head.Font.Name = "Times New Roman";

                head.Font.Size = "20";
                head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ////
                int nam = Convert.ToInt32(cbo_namgv.Text);
                string query = " select MaGV,HoTen,NgaySinh,KhoaChuQuan,DonViCongTac from GiangVien where MaGV = '"+txt_magv.Text+"' ";
                DataTable tb = my.DocDL(query);

                Excel.Range head1 = oSheet.get_Range("A2", "H2");

                head1.MergeCells = true;
                string ngay = DateTime.Now.ToString("dd-MM-yyyy");
                head1.Value2 = "MaGV : " + tb.Rows[0]["MaGV"].ToString();

                head1.Font.Bold = true;

                head1.Font.Name = "Times New Roman";

                head1.Font.Size = "13";

                

                /////

                Excel.Range head2 = oSheet.get_Range("A3", "G3");

                head2.MergeCells = true;

                head2.Value2 = "Họ Tên : " + tb.Rows[0]["HoTen"].ToString();

                head2.Font.Bold = true;

                head2.Font.Name = "Times New Roman";

                head2.Font.Size = "13";

                


                /////


                Excel.Range head3 = oSheet.get_Range("A4", "G4");

                head3.MergeCells = true;
                string ngaysinh = ((DateTime)tb.Rows[0]["NgaySinh"]).ToString("dd-MM-yyyy");
                head3.Value2 = "Ngày sinh : " + ngaysinh;

                head3.Font.Bold = true;

                head3.Font.Name = "Times New Roman";

                head3.Font.Size = "13";

                


                /////
                Excel.Range head4 = oSheet.get_Range("A5", "G5");

                head4.MergeCells = true;

                head4.Value2 = "Khoa : " + tb.Rows[0]["KhoaChuQuan"].ToString();

                head4.Font.Bold = true;

                head4.Font.Name = "Times New Roman";

                head4.Font.Size = "13";




                /////

                Excel.Range head5 = oSheet.get_Range("A6", "G6");

                head5.MergeCells = true;

                head5.Value2 = "Đơn vị công tác : " + tb.Rows[0]["DonViCongTac"].ToString();

                head5.Font.Bold = true;

                head5.Font.Name = "Times New Roman";

                head5.Font.Size = "13";




                /////
                Excel.Range cl1 = oSheet.get_Range("A7", "A7");
                cl1.Value = "Mã đề tài";

                Excel.Range cl2 = oSheet.get_Range("B7", "B7");
                cl2.Value = "Tên đề tài";

                Excel.Range cl3 = oSheet.get_Range("C7", "C7");
                cl3.Value = "Lĩnh vực";

                Excel.Range cl4 = oSheet.get_Range("D7", "D7");
                cl4.Value = "Khoa";

                Excel.Range cl5 = oSheet.get_Range("E7", "E7");
                cl5.Value = "Ngày bắt đầu";

                Excel.Range cl10 = oSheet.get_Range("F7", "F7");
                cl10.Value = "Ngày kết thúc";

                Excel.Range cl6 = oSheet.get_Range("G7", "G7");
                cl6.Value = "Tiến độ";

                Excel.Range cl7 = oSheet.get_Range("H7", "H7");
                cl7.Value = "Cấp đề tài";

                //Excel.Range cl8 = oSheet.get_Range("I3", "I3");
                //cl8.Value = "";

                //Excel.Range cl9 = oSheet.get_Range("J3", "J3");
                //cl9.Value = "Giảng viên hướng dẫn";





                Excel.Range rowHead = oSheet.get_Range("A7", "H7");
                rowHead.Font.Bold = true;
                rowHead.Font.Name = "Times New Roman";
                rowHead.Font.Size = 13;
                rowHead.Borders.LineStyle = Excel.Constants.xlSolid;
                rowHead.Interior.ColorIndex = 6;
                rowHead.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Sau đó, thêm dữ liệu từ DataTable
                int line = 8;


                for (int i = 0; i < dataTable.Rows.Count; i++)
                {

                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        oSheet.Cells[i + line, j + 1] = dataTable.Rows[i][j];
                        oSheet.Cells[i + line, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        oSheet.Cells[i + line, j + 1].Borders.LineStyle = Excel.Constants.xlSolid;
                        oSheet.Cells[i + line, j + 1].Font.Name = "Times New Roman";

                    }

                }
                ////
                ///
                /// 
                /// 
                int dong = line + dataTable.Rows.Count;
                Excel.Range headSum = oSheet.get_Range("A" + dong, "H" + dong);

                headSum.MergeCells = true;

                headSum.Value2 = "Tổng đề tài : " + dataTable.Rows.Count;

                headSum.Font.Bold = true;

                headSum.Font.Name = "Times New Roman";

                headSum.Font.Size = "13";
                headSum.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                headSum.Borders.LineStyle = Excel.Constants.xlSolid;
                //////
                oSheet.Name = "BCDTGV";
                oExcel.Columns.AutoFit();

                workbook.Activate();
                SaveFileDialog saveFile = new SaveFileDialog();
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    saveFile.Filter = "Text Files|*.xlxs|All Files|*.*";
                    workbook.SaveAs(saveFile.FileName.ToLower());
                    MessageBox.Show("Xuất danh sách thành công", "Thông báo");
                }

                oExcel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi xuất báo cáo: {ex.Message}");
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txt_magv.Text) || string.IsNullOrWhiteSpace(cbo_namgv.Text))
            {
                MessageBox.Show("Vui lòng chọn thông tin xuất báo cáo", "Thông báo");
            }
            else
            {
                excelGVDT();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            

        }
    }
}