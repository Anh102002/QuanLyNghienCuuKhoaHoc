
namespace QL_NCKH
{
    partial class uc_sinhvien
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_sinhvien));
            this.barManager1 = new DevExpress.XtraBars.BarManager(this.components);
            this.bar2 = new DevExpress.XtraBars.Bar();
            this.btn_them = new DevExpress.XtraBars.BarButtonItem();
            this.btn_sua = new DevExpress.XtraBars.BarButtonItem();
            this.btn_xoa = new DevExpress.XtraBars.BarButtonItem();
            this.btn_export = new DevExpress.XtraBars.BarButtonItem();
            this.btn_refesh = new DevExpress.XtraBars.BarButtonItem();
            this.btn_giayto = new DevExpress.XtraBars.BarButtonItem();
            this.barDockControlTop = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlBottom = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlLeft = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlRight = new DevExpress.XtraBars.BarDockControl();
            this.btn_timkiem = new DevExpress.XtraBars.BarButtonItem();
            this.btn_thoat = new DevExpress.XtraBars.BarButtonItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.txt_timkiem = new DevExpress.XtraEditors.SearchControl();
            this.cb_khoa = new System.Windows.Forms.ComboBox();
            this.dtp_ngaysinh = new System.Windows.Forms.DateTimePicker();
            this.cb_gioitinh = new System.Windows.Forms.ComboBox();
            this.cb_coso = new System.Windows.Forms.ComboBox();
            this.cbo_tk = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txt_sodt = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txt_email = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txt_lop = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_hoten = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_masv = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dgv_sinhvien = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.barManager1)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txt_timkiem.Properties)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_sinhvien)).BeginInit();
            this.SuspendLayout();
            // 
            // barManager1
            // 
            this.barManager1.Bars.AddRange(new DevExpress.XtraBars.Bar[] {
            this.bar2});
            this.barManager1.DockControls.Add(this.barDockControlTop);
            this.barManager1.DockControls.Add(this.barDockControlBottom);
            this.barManager1.DockControls.Add(this.barDockControlLeft);
            this.barManager1.DockControls.Add(this.barDockControlRight);
            this.barManager1.Form = this;
            this.barManager1.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.btn_them,
            this.btn_sua,
            this.btn_xoa,
            this.btn_timkiem,
            this.btn_thoat,
            this.btn_export,
            this.btn_refesh,
            this.btn_giayto});
            this.barManager1.MaxItemId = 8;
            // 
            // bar2
            // 
            this.bar2.BarName = "Tools";
            this.bar2.DockCol = 0;
            this.bar2.DockRow = 0;
            this.bar2.DockStyle = DevExpress.XtraBars.BarDockStyle.Top;
            this.bar2.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
            new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.btn_them, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph),
            new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.btn_sua, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph),
            new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.btn_xoa, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph),
            new DevExpress.XtraBars.LinkPersistInfo(this.btn_export, true),
            new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.btn_refesh, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph),
            new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.btn_giayto, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph)});
            this.bar2.Text = "Tools";
            // 
            // btn_them
            // 
            this.btn_them.Caption = "Thêm\r\n";
            this.btn_them.Id = 0;
            this.btn_them.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btn_them.ImageOptions.Image")));
            this.btn_them.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btn_them.ImageOptions.LargeImage")));
            this.btn_them.Name = "btn_them";
            this.btn_them.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btn_them_ItemClick);
            // 
            // btn_sua
            // 
            this.btn_sua.Caption = "Sửa";
            this.btn_sua.Id = 1;
            this.btn_sua.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("btn_sua.ImageOptions.SvgImage")));
            this.btn_sua.Name = "btn_sua";
            this.btn_sua.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btn_sua_ItemClick);
            // 
            // btn_xoa
            // 
            this.btn_xoa.Caption = "Xóa";
            this.btn_xoa.Id = 2;
            this.btn_xoa.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("btn_xoa.ImageOptions.SvgImage")));
            this.btn_xoa.Name = "btn_xoa";
            this.btn_xoa.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btn_xoa_ItemClick);
            // 
            // btn_export
            // 
            this.btn_export.Caption = "Export danh sách";
            this.btn_export.Id = 5;
            this.btn_export.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btn_export.ImageOptions.Image")));
            this.btn_export.Name = "btn_export";
            this.btn_export.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph;
            this.btn_export.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btn_export_ItemClick);
            // 
            // btn_refesh
            // 
            this.btn_refesh.Caption = "Refresh";
            this.btn_refesh.Id = 6;
            this.btn_refesh.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btn_refesh.ImageOptions.Image")));
            this.btn_refesh.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btn_refesh.ImageOptions.LargeImage")));
            this.btn_refesh.Name = "btn_refesh";
            this.btn_refesh.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btn_refesh_ItemClick);
            // 
            // btn_giayto
            // 
            this.btn_giayto.Caption = "Giấy tờ liên quan";
            this.btn_giayto.Id = 7;
            this.btn_giayto.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btn_giayto.ImageOptions.Image")));
            this.btn_giayto.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btn_giayto.ImageOptions.LargeImage")));
            this.btn_giayto.Name = "btn_giayto";
            this.btn_giayto.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btn_giayto_ItemClick);
            // 
            // barDockControlTop
            // 
            this.barDockControlTop.Appearance.Font = new System.Drawing.Font("Century Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.barDockControlTop.Appearance.Options.UseFont = true;
            this.barDockControlTop.CausesValidation = false;
            this.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.barDockControlTop.Location = new System.Drawing.Point(0, 0);
            this.barDockControlTop.Manager = this.barManager1;
            this.barDockControlTop.Size = new System.Drawing.Size(817, 24);
            // 
            // barDockControlBottom
            // 
            this.barDockControlBottom.CausesValidation = false;
            this.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.barDockControlBottom.Location = new System.Drawing.Point(0, 515);
            this.barDockControlBottom.Manager = this.barManager1;
            this.barDockControlBottom.Size = new System.Drawing.Size(817, 0);
            // 
            // barDockControlLeft
            // 
            this.barDockControlLeft.CausesValidation = false;
            this.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left;
            this.barDockControlLeft.Location = new System.Drawing.Point(0, 24);
            this.barDockControlLeft.Manager = this.barManager1;
            this.barDockControlLeft.Size = new System.Drawing.Size(0, 491);
            // 
            // barDockControlRight
            // 
            this.barDockControlRight.CausesValidation = false;
            this.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right;
            this.barDockControlRight.Location = new System.Drawing.Point(817, 24);
            this.barDockControlRight.Manager = this.barManager1;
            this.barDockControlRight.Size = new System.Drawing.Size(0, 491);
            // 
            // btn_timkiem
            // 
            this.btn_timkiem.Caption = "Tìm kiếm";
            this.btn_timkiem.Id = 3;
            this.btn_timkiem.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("btn_timkiem.ImageOptions.SvgImage")));
            this.btn_timkiem.Name = "btn_timkiem";
            this.btn_timkiem.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btn_timkiem_ItemClick);
            // 
            // btn_thoat
            // 
            this.btn_thoat.Caption = "Thoát";
            this.btn_thoat.Id = 4;
            this.btn_thoat.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btn_thoat.ImageOptions.Image")));
            this.btn_thoat.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btn_thoat.ImageOptions.LargeImage")));
            this.btn_thoat.Name = "btn_thoat";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.txt_timkiem);
            this.groupBox1.Controls.Add(this.cb_khoa);
            this.groupBox1.Controls.Add(this.dtp_ngaysinh);
            this.groupBox1.Controls.Add(this.cb_gioitinh);
            this.groupBox1.Controls.Add(this.cb_coso);
            this.groupBox1.Controls.Add(this.cbo_tk);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.txt_sodt);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.txt_email);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.txt_lop);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txt_hoten);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.txt_masv);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(0, 24);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(817, 234);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Thông tin sinh viên";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.button1.Image = global::QL_NCKH.Properties.Resources.search_16;
            this.button1.Location = new System.Drawing.Point(726, 194);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 91;
            this.button1.Text = "Tìm kiếm";
            this.button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txt_timkiem
            // 
            this.txt_timkiem.Location = new System.Drawing.Point(282, 191);
            this.txt_timkiem.MenuManager = this.barManager1;
            this.txt_timkiem.Name = "txt_timkiem";
            this.txt_timkiem.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Repository.ClearButton()});
            this.txt_timkiem.Properties.ShowSearchButton = false;
            this.txt_timkiem.Size = new System.Drawing.Size(438, 20);
            this.txt_timkiem.TabIndex = 90;
            this.txt_timkiem.SelectedIndexChanged += new System.EventHandler(this.txt_timkiem_SelectedIndexChanged);
            // 
            // cb_khoa
            // 
            this.cb_khoa.BackColor = System.Drawing.Color.White;
            this.cb_khoa.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_khoa.FormattingEnabled = true;
            this.cb_khoa.Items.AddRange(new object[] {
            "Công nghệ thông tin",
            "Điện tử",
            "Cơ khí",
            "Kế toán Kiểm toán",
            "Ngoại Ngữ",
            "Du lịch và khách sạn",
            "Tài chình ngân hàng và BH",
            "Thương mại",
            "Điện - Tự động hóa",
            "Diệt may và Thời trang",
            " Lý luận chính trị và Pháp luật",
            "Quản trị kinh doanh",
            "Công nghệ thực phẩm",
            "Quản trị & Marketing",
            "Giáo dục thể chất",
            "Kế toán Kiểm toán",
            "Khoa học ứng dụng"});
            this.cb_khoa.Location = new System.Drawing.Point(315, 106);
            this.cb_khoa.Name = "cb_khoa";
            this.cb_khoa.Size = new System.Drawing.Size(165, 21);
            this.cb_khoa.TabIndex = 33;
            // 
            // dtp_ngaysinh
            // 
            this.dtp_ngaysinh.CustomFormat = "dd/MM/yyyy";
            this.dtp_ngaysinh.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtp_ngaysinh.Location = new System.Drawing.Point(60, 158);
            this.dtp_ngaysinh.Name = "dtp_ngaysinh";
            this.dtp_ngaysinh.Size = new System.Drawing.Size(165, 20);
            this.dtp_ngaysinh.TabIndex = 32;
            // 
            // cb_gioitinh
            // 
            this.cb_gioitinh.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_gioitinh.FormattingEnabled = true;
            this.cb_gioitinh.Items.AddRange(new object[] {
            "Nam",
            "Nữ"});
            this.cb_gioitinh.Location = new System.Drawing.Point(568, 158);
            this.cb_gioitinh.Name = "cb_gioitinh";
            this.cb_gioitinh.Size = new System.Drawing.Size(165, 21);
            this.cb_gioitinh.TabIndex = 31;
            // 
            // cb_coso
            // 
            this.cb_coso.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_coso.FormattingEnabled = true;
            this.cb_coso.Items.AddRange(new object[] {
            "Hà Nội",
            "Nam Định"});
            this.cb_coso.Location = new System.Drawing.Point(568, 106);
            this.cb_coso.Name = "cb_coso";
            this.cb_coso.Size = new System.Drawing.Size(165, 21);
            this.cb_coso.TabIndex = 30;
            // 
            // cbo_tk
            // 
            this.cbo_tk.BackColor = System.Drawing.Color.White;
            this.cbo_tk.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbo_tk.FormattingEnabled = true;
            this.cbo_tk.Items.AddRange(new object[] {
            "Mã sinh viên",
            "Họ tên",
            "Lớp"});
            this.cbo_tk.Location = new System.Drawing.Point(139, 195);
            this.cbo_tk.Name = "cbo_tk";
            this.cbo_tk.Size = new System.Drawing.Size(121, 21);
            this.cbo_tk.TabIndex = 19;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(60, 198);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(73, 13);
            this.label10.TabIndex = 18;
            this.label10.Text = "Tìm kiếm theo";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(568, 136);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(51, 13);
            this.label7.TabIndex = 16;
            this.label7.Text = "Giới Tính";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(568, 84);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(36, 13);
            this.label8.TabIndex = 14;
            this.label8.Text = "Cơ Sở";
            // 
            // txt_sodt
            // 
            this.txt_sodt.Location = new System.Drawing.Point(568, 54);
            this.txt_sodt.Name = "txt_sodt";
            this.txt_sodt.Size = new System.Drawing.Size(165, 20);
            this.txt_sodt.TabIndex = 13;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(568, 32);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(75, 13);
            this.label9.TabIndex = 12;
            this.label9.Text = "Số Điện Thoại";
            // 
            // txt_email
            // 
            this.txt_email.Location = new System.Drawing.Point(315, 158);
            this.txt_email.Name = "txt_email";
            this.txt_email.Size = new System.Drawing.Size(165, 20);
            this.txt_email.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(315, 136);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(32, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Email";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(315, 84);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(32, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Khoa";
            // 
            // txt_lop
            // 
            this.txt_lop.Location = new System.Drawing.Point(315, 54);
            this.txt_lop.Name = "txt_lop";
            this.txt_lop.Size = new System.Drawing.Size(165, 20);
            this.txt_lop.TabIndex = 7;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(315, 32);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(25, 13);
            this.label6.TabIndex = 6;
            this.label6.Text = "Lớp";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(60, 136);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Ngày Sinh";
            // 
            // txt_hoten
            // 
            this.txt_hoten.Location = new System.Drawing.Point(60, 106);
            this.txt_hoten.Name = "txt_hoten";
            this.txt_hoten.Size = new System.Drawing.Size(165, 20);
            this.txt_hoten.TabIndex = 3;
            this.txt_hoten.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(60, 84);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Họ Tên";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // txt_masv
            // 
            this.txt_masv.Location = new System.Drawing.Point(60, 54);
            this.txt_masv.Name = "txt_masv";
            this.txt_masv.Size = new System.Drawing.Size(165, 20);
            this.txt_masv.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(60, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Mã sinh viên";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dgv_sinhvien);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(0, 258);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(817, 257);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Danh sách sinh viên";
            // 
            // dgv_sinhvien
            // 
            this.dgv_sinhvien.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_sinhvien.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_sinhvien.Location = new System.Drawing.Point(3, 16);
            this.dgv_sinhvien.Name = "dgv_sinhvien";
            this.dgv_sinhvien.ReadOnly = true;
            this.dgv_sinhvien.Size = new System.Drawing.Size(811, 238);
            this.dgv_sinhvien.TabIndex = 0;
            this.dgv_sinhvien.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_sinhvien_CellClick);
            this.dgv_sinhvien.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_sinhvien_CellContentClick);
            // 
            // uc_sinhvien
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.barDockControlLeft);
            this.Controls.Add(this.barDockControlRight);
            this.Controls.Add(this.barDockControlBottom);
            this.Controls.Add(this.barDockControlTop);
            this.Name = "uc_sinhvien";
            this.Size = new System.Drawing.Size(817, 515);
            this.Load += new System.EventHandler(this.uc_sinhvien_Load);
            ((System.ComponentModel.ISupportInitialize)(this.barManager1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txt_timkiem.Properties)).EndInit();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_sinhvien)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraBars.BarManager barManager1;
        private DevExpress.XtraBars.Bar bar2;
        private DevExpress.XtraBars.BarButtonItem btn_them;
        private DevExpress.XtraBars.BarButtonItem btn_sua;
        private DevExpress.XtraBars.BarButtonItem btn_xoa;
        private DevExpress.XtraBars.BarButtonItem btn_timkiem;
        private DevExpress.XtraBars.BarButtonItem btn_export;
        private DevExpress.XtraBars.BarDockControl barDockControlTop;
        private DevExpress.XtraBars.BarDockControl barDockControlBottom;
        private DevExpress.XtraBars.BarDockControl barDockControlLeft;
        private DevExpress.XtraBars.BarDockControl barDockControlRight;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView dgv_sinhvien;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txt_hoten;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_masv;
        private System.Windows.Forms.Label label1;
        private DevExpress.XtraBars.BarButtonItem btn_thoat;
        private System.Windows.Forms.ComboBox cbo_tk;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txt_sodt;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txt_email;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txt_lop;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cb_coso;
        private System.Windows.Forms.ComboBox cb_gioitinh;
        private System.Windows.Forms.DateTimePicker dtp_ngaysinh;
        private DevExpress.XtraBars.BarButtonItem btn_refesh;
        private DevExpress.XtraBars.BarButtonItem btn_giayto;
        private System.Windows.Forms.ComboBox cb_khoa;
        private System.Windows.Forms.Button button1;
        private DevExpress.XtraEditors.SearchControl txt_timkiem;
    }
}
