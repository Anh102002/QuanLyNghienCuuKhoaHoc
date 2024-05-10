using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Linq;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraCharts;

namespace QL_NCKH
{
    public partial class uc_thongke : DevExpress.XtraEditors.XtraUserControl
    {
        MyClass my = new MyClass();
        
        
        public uc_thongke()
        {
            InitializeComponent();
        }
        public void loadDL()
        {
            try
            {
                string ngaykt = DateTime.Now.ToString("yyyy-MM-dd");
                string sql = "select  * from dbo.SoLuongDeTaiTrongNam('2010-01-01','" + ngaykt + "',N'Giảng viên') ";
                DataTable tb = my.DocDL(sql);

                DataGridView dataGridView1 = new DataGridView();

                dataGridView1.DataSource = tb;
                



                chartControl1.Series.Clear();
                chartControl1.Titles.Clear();

                Color[] colors = { Color.Red, Color.Blue, Color.Green, Color.Orange, Color.Purple, Color.Yellow, Color.Cyan, Color.Magenta, Color.Lime, Color.Teal };

                // Index của màu tiếp theo trong danh sách
                int colorIndex = 0;

                // Tạo một series cho mỗi cấp độ
                if(tb.Rows.Count > 0)
                {
                    foreach (DataRow row in tb.Rows)
                    {
                        string capDeTai = row["CapDeTai"].ToString();
                        Series series = new Series(capDeTai, ViewType.Bar);
                        series.Points.Add(new SeriesPoint(Convert.ToInt32(row["Nam"]), Convert.ToInt32(row["SoDeTai"])));

                        // Thiết lập màu sắc cho series
                        series.View.Color = colors[colorIndex];

                        // Tăng index màu sắc
                        colorIndex = (colorIndex + 1) % colors.Length;

                        // Thêm series vào biểu đồ
                        chartControl1.Series.Add(series);

                        

                    }

               // Thiết lập tiêu đề cho trục x và y
                    ((XYDiagram)chartControl1.Diagram).AxisX.Title.Text = "Năm";
                    ((XYDiagram)chartControl1.Diagram).AxisY.Title.Text = "Số đề tài";
                    ((XYDiagram)chartControl1.Diagram).AxisX.Label.TextPattern = "{V:0}";
                    ((XYDiagram)chartControl1.Diagram).AxisX.WholeRange.Auto = false;
                    ((XYDiagram)chartControl1.Diagram).AxisX.WholeRange.SetMinMaxValues(Convert.ToInt32(tb.Compute("MIN(Nam)", "")), Convert.ToInt32(tb.Compute("MAX(Nam)", "")));

                    ((XYDiagram)chartControl1.Diagram).AxisX.GridSpacingAuto = false;
                    ((XYDiagram)chartControl1.Diagram).AxisX.GridSpacing = 1;

                    ((XYDiagram)chartControl1.Diagram).AxisY.GridSpacingAuto = false;
                    ((XYDiagram)chartControl1.Diagram).AxisY.GridSpacing = 1;

                    ChartTitle chartTitle = new ChartTitle();
                    chartTitle.Text = "Biểu đồ số lượng đề tài \ntheo cấp độ và năm";
                    chartTitle.Font = new Font("Tahoma", 10);
                    chartControl1.Titles.Add(chartTitle);
                    
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi {"+ex.Message+"}");
            }
        }
        private void uc_thongke_Load(object sender, EventArgs e)
        {

            loadDL();
            loadDLSV();
            LoadDLBB();

            string nam = DateTime.Now.Year.ToString();
            cbo_namht.Text = nam;
            cbo_namkp.Text = nam;
            
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            loadDL();
            loadDLSV();
            LoadDLBB();
            txt_timkiem.Clear();
            cbo_nam.SelectedIndex = -1;
            chartControl2.Series.Clear();
            string nam = DateTime.Now.Year.ToString();
            cbo_namht.Text = nam;
            cbo_namkp.Text = nam;
        }
        public void loadDLGV()
        {
            try
            {
                
                string sql = "select  * from dbo.SoLuongThanhVienDT('"+txt_timkiem.Text+"','" + cbo_nam.Text + "') ";
                DataTable tb = my.DocDL(sql);

                
                



                chartControl2.Series.Clear();
                chartControl2.Titles.Clear();

                //Color[] colors = { Color.Red, Color.Blue, Color.Green, Color.Orange, Color.Purple, Color.Yellow, Color.Cyan, Color.Magenta, Color.Lime, Color.Teal };

                // Index của màu tiếp theo trong danh sách
                int colorIndex = 0;

                // Tạo một series cho mỗi cấp độ
                if (tb.Rows.Count > 0)
                {
                    foreach (DataRow row in tb.Rows)
                    {
                        string capDeTai = row["Magiangvien"].ToString();
                        Series series = new Series(capDeTai, ViewType.Bar);
                        series.Points.Add(new SeriesPoint(Convert.ToInt32(row["Nam"]), Convert.ToInt32(row["SoDeTai"])));

                        // Thiết lập màu sắc cho series
                        //series.View.Color = colors[colorIndex];

                        // Tăng index màu sắc
                        //colorIndex = (colorIndex + 1) % colors.Length;

                        // Thêm series vào biểu đồ
                        chartControl2.Series.Add(series);


                    }

               // Thiết lập tiêu đề cho trục x và y
                    ((XYDiagram)chartControl2.Diagram).AxisX.Title.Text = "Năm";
                    ((XYDiagram)chartControl2.Diagram).AxisY.Title.Text = "Số đề tài";
                    ((XYDiagram)chartControl2.Diagram).AxisX.Label.TextPattern = "{V:0}";
                    ((XYDiagram)chartControl2.Diagram).AxisX.WholeRange.Auto = false;
                    ((XYDiagram)chartControl2.Diagram).AxisX.WholeRange.SetMinMaxValues(Convert.ToInt32(tb.Compute("MIN(Nam)", "")), Convert.ToInt32(tb.Compute("MAX(Nam)", "")));

                    ((XYDiagram)chartControl2.Diagram).AxisX.GridSpacingAuto = false;
                    ((XYDiagram)chartControl2.Diagram).AxisX.GridSpacing = 1;

                    ((XYDiagram)chartControl2.Diagram).AxisY.GridSpacingAuto = false;
                    ((XYDiagram)chartControl2.Diagram).AxisY.GridSpacing = 1;



                    //ChartTitle chartTitle = new ChartTitle();
                    //chartTitle.Text = "Biểu đồ số lượng đề tài \ngiảng viên tham gia theo năm";
                    //chartTitle.Font = new Font("Tahoma", 10);
                    //chartControl2.Titles.Add(chartTitle);
                }
                  
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi {" + ex.Message + "}");
            }
        }


        public void loadDLSV()
        {
            try
            {

                string ngaykt = DateTime.Now.ToString("yyyy-MM-dd");
                string sql = "SELECT * from dbo.SoLuongSVDT('2010-01-01','"+ngaykt+"',N'Sinh viên')  ";
                DataTable tb = my.DocDL(sql);

                chartControl3.Series.Clear();
                chartControl3.Titles.Clear();

                if(tb.Rows.Count > 0)
                {
                    Series series1 = new Series("Số sinh viên", ViewType.Bar);
                    foreach (DataRow row in tb.Rows)
                    {
                        series1.Points.Add(new SeriesPoint(Convert.ToInt32(row["Nam"]), Convert.ToInt32(row["SoSV"])));
                    }

                    // Tạo series cho số đề tài
                    Series series2 = new Series("Số đề tài", ViewType.Bar);
                    foreach (DataRow row in tb.Rows)
                    {
                        series2.Points.Add(new SeriesPoint(Convert.ToInt32(row["Nam"]), Convert.ToInt32(row["SoDeTai"])));
                    }

                    // Thêm series vào biểu đồ
                    chartControl3.Series.AddRange(new Series[] { series1, series2 });

                    // Thiết lập tiêu đề cho trục x và y
                    ((XYDiagram)chartControl3.Diagram).AxisX.Title.Text = "Năm";
                    ((XYDiagram)chartControl3.Diagram).AxisY.Title.Text = "Số lượng";
                    ((XYDiagram)chartControl3.Diagram).AxisX.Label.TextPattern = "{V:0}";
                    ((XYDiagram)chartControl3.Diagram).AxisX.WholeRange.Auto = false;
                    ((XYDiagram)chartControl3.Diagram).AxisX.WholeRange.SetMinMaxValues(Convert.ToInt32(tb.Compute("MIN(Nam)", "")), Convert.ToInt32(tb.Compute("MAX(Nam)", "")));

                    ((XYDiagram)chartControl3.Diagram).AxisX.GridSpacingAuto = false;
                    ((XYDiagram)chartControl3.Diagram).AxisX.GridSpacing = 1;

                    ((XYDiagram)chartControl3.Diagram).AxisY.GridSpacingAuto = false;
                    ((XYDiagram)chartControl3.Diagram).AxisY.GridSpacing = 1;


                    ChartTitle chartTitle = new ChartTitle();
                    chartTitle.Text = "Biểu đồ số lượng sinh viên tham gia đề tài";
                    chartTitle.Font = new Font("Tahoma", 10);
                    chartControl3.Titles.Add(chartTitle);


                }




            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi {" + ex.Message + "}");
            }
        }

        public void LoadDLBB()
        {
            try
            {
                string bbgt = "select MaBB from BaiBao where LoaiBaiBao = N'Quốc tế' ";
                DataTable tb = my.DocDL(bbgt);

                string bbtn = "select MaBB from BaiBao where LoaiBaiBao = N'Trong nước' ";
                DataTable dt = my.DocDL(bbtn);

                string tc = "select MaTC from TapChi";
                DataTable table = my.DocDL(tc);

                

                chartControl4.Series.Clear();
                chartControl4.Titles.Clear();
                Series series1 = new Series("Số bài báo", ViewType.Bar);
                if (tb.Rows.Count > 0)
                {
                    
                    series1.Points.Add(new SeriesPoint("Bài báo quốc tế", tb.Rows.Count));
                }
                if(dt.Rows.Count > 0)
                {
                    series1.Points.Add(new SeriesPoint("Bài báo trong nước", dt.Rows.Count));
                }
                Series series2 = new Series("Số tập chí", ViewType.Bar);
                if (table.Rows.Count > 0)
                {
                    
                    series2.Points.Add(new SeriesPoint("Tạp chí NCKH", table.Rows.Count));
                }

                

                

                chartControl4.Series.Add(series1);
                chartControl4.Series.Add(series2);

                ((XYDiagram)chartControl4.Diagram).AxisY.GridSpacingAuto = false;
                ((XYDiagram)chartControl4.Diagram).AxisY.GridSpacing = 1;

                ChartTitle chartTitle = new ChartTitle();
                chartTitle.Text = "Biểu đồ số lượng bài báo và tạp chí được xuất bản";
                chartTitle.Font = new Font("Tahoma", 10);
                chartControl4.Titles.Add(chartTitle);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi {" + ex.Message + "}");
            }
        }

        public void LoadDLHT()
        {
            try
            {
                string sql = "SELECT * FROM dbo.SoLuongHTtrongNam('"+cbo_namht.Text+"') ";
                DataTable dataTable = my.DocDL(sql);

                chartControl5.Series.Clear();
                chartControl5.Titles.Clear();


                if (dataTable.Rows.Count > 0)
                {
                    Series series = new Series("Số hội thảo", ViewType.Pie);

                    // Thêm dữ liệu vào biểu đồ tròn
                    foreach (DataRow row in dataTable.Rows)
                    {
                        string capHoiThao = row["CapHoiThao"].ToString();
                        int soDeTai = Convert.ToInt32(row["SoDeTai"]);
                        series.Points.Add(new SeriesPoint(capHoiThao, soDeTai));
                    }

                    // Thêm series vào biểu đồ
                    chartControl5.Series.Add(series);

                    PieSeriesLabel label = (PieSeriesLabel)series.Label;
                    label.Position = PieSeriesLabelPosition.TwoColumns;
                    label.TextPattern = "{A}: {VP:0.##%}";

                    ChartTitle chartTitle = new ChartTitle();
                    chartTitle.Text = "Biểu đồ tỷ lệ hội thảo các cấp theo năm";
                    chartTitle.Font = new Font("Tahoma", 10);
                    chartControl5.Titles.Add(chartTitle);
                }
                

                

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi {" + ex.Message + "}");
            }
        }


        public void LoadKP()
        {
            try
            {
                int nam = Convert.ToInt32(cbo_namkp.Text);
                string sql = "SELECT * FROM TongKinhPhiTheoNam('"+nam+"',N'Giảng viên') ";
                DataTable tb = my.DocDL(sql);

                chartControl6.Series.Clear();
                chartControl6.Titles.Clear();
                if(tb.Rows.Count > 0)
                {
                    Color[] colors = { Color.Red, Color.Blue, Color.Green, Color.Orange, Color.Purple, Color.Yellow, Color.Cyan, Color.Magenta, Color.Lime, Color.Teal };

                    // Index của màu tiếp theo trong danh sách
                    int colorIndex = 0;

                    // Tạo một series cho mỗi cấp độ
                    if (tb.Rows.Count > 0)
                    {
                        foreach (DataRow row in tb.Rows)
                        {
                            string capDeTai = row["CapDeTai"].ToString();
                            Series series = new Series(capDeTai, ViewType.Bar);
                            series.Points.Add(new SeriesPoint(Convert.ToInt32(row["Nam"]), Convert.ToInt32(row["TongKinhPhi"])));

                            // Thiết lập màu sắc cho series
                            series.View.Color = colors[colorIndex];

                            // Tăng index màu sắc
                            colorIndex = (colorIndex + 1) % colors.Length;

                            // Thêm series vào biểu đồ
                            chartControl6.Series.Add(series);



                        }

                        // Thiết lập tiêu đề cho trục x và y
                        ((XYDiagram)chartControl6.Diagram).AxisX.Title.Text = "Năm";
                        ((XYDiagram)chartControl6.Diagram).AxisY.Title.Text = "Kinh phí(VND)";
                        ((XYDiagram)chartControl6.Diagram).AxisX.Label.TextPattern = "{V:0}";
                        ((XYDiagram)chartControl6.Diagram).AxisX.WholeRange.Auto = false;
                        ((XYDiagram)chartControl6.Diagram).AxisX.WholeRange.SetMinMaxValues(Convert.ToInt32(tb.Compute("MIN(Nam)", "")), Convert.ToInt32(tb.Compute("MAX(Nam)", "")));



                        ((XYDiagram)chartControl6.Diagram).AxisX.GridSpacingAuto = false;
                        ((XYDiagram)chartControl6.Diagram).AxisX.GridSpacing = 1;

                        ((XYDiagram)chartControl6.Diagram).AxisY.GridSpacingAuto = false;
                        ((XYDiagram)chartControl6.Diagram).AxisY.GridSpacing = 1;

                        ChartTitle chartTitle = new ChartTitle();
                        chartTitle.Text = "Biểu đồ kinh phí đề tài các cấp theo năm";
                        chartTitle.Font = new Font("Tahoma", 10);
                        chartControl6.Titles.Add(chartTitle);

                    }
                }
                






                

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi {" + ex.Message + "}");
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrWhiteSpace(txt_timkiem.Text) || string.IsNullOrWhiteSpace(cbo_nam.Text))
            {
                MessageBox.Show("Vui lòng nhập đầy đủ thông tin", "Thông báo");
            }
            else
            {
                loadDLGV();
            }
        }

        private void cbo_namht_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadDLHT();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadKP();
        }

        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
           
        }

        private void barButtonItem7_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            frm_thongketheonam frm = new frm_thongketheonam();
            frm.ShowDialog();
        }
    }
}
