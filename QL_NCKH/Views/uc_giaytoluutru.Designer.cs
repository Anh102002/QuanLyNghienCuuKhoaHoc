namespace QL_NCKH
{
    partial class uc_giaytoluutru
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_giaytoluutru));
            this.barManager1 = new DevExpress.XtraBars.BarManager(this.components);
            this.bar2 = new DevExpress.XtraBars.Bar();
            this.btn_them = new DevExpress.XtraBars.BarButtonItem();
            this.btn_sua = new DevExpress.XtraBars.BarButtonItem();
            this.btn_xoa = new DevExpress.XtraBars.BarButtonItem();
            this.btn_openfile = new DevExpress.XtraBars.BarButtonItem();
            this.barDockControlTop = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlBottom = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlLeft = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlRight = new DevExpress.XtraBars.BarDockControl();
            this.btn_timkiem = new DevExpress.XtraBars.BarButtonItem();
            this.btn_thoat = new DevExpress.XtraBars.BarButtonItem();
            this.btn_export = new DevExpress.XtraBars.BarButtonItem();
            this.btn_refresh = new DevExpress.XtraBars.BarButtonItem();
            this.barButtonItem1 = new DevExpress.XtraBars.BarButtonItem();
            this.btn_giayto = new DevExpress.XtraBars.BarButtonItem();
            this.barButtonItem2 = new DevExpress.XtraBars.BarButtonItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txt_magt = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cbo_luutru = new System.Windows.Forms.ComboBox();
            this.btn__choose = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_url = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_tengt = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dgv_giayto = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.barManager1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_giayto)).BeginInit();
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
            this.btn_refresh,
            this.barButtonItem1,
            this.btn_giayto,
            this.btn_openfile,
            this.barButtonItem2});
            this.barManager1.MaxItemId = 11;
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
            new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.btn_openfile, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph)});
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
            // btn_openfile
            // 
            this.btn_openfile.Caption = "Mở File";
            this.btn_openfile.Id = 9;
            this.btn_openfile.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btn_openfile.ImageOptions.Image")));
            this.btn_openfile.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btn_openfile.ImageOptions.LargeImage")));
            this.btn_openfile.Name = "btn_openfile";
            this.btn_openfile.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btn_openfile_ItemClick);
            // 
            // barDockControlTop
            // 
            this.barDockControlTop.Appearance.Font = new System.Drawing.Font("Century Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.barDockControlTop.Appearance.Options.UseFont = true;
            this.barDockControlTop.CausesValidation = false;
            this.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.barDockControlTop.Location = new System.Drawing.Point(0, 0);
            this.barDockControlTop.Manager = this.barManager1;
            this.barDockControlTop.Size = new System.Drawing.Size(572, 24);
            // 
            // barDockControlBottom
            // 
            this.barDockControlBottom.CausesValidation = false;
            this.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.barDockControlBottom.Location = new System.Drawing.Point(0, 447);
            this.barDockControlBottom.Manager = this.barManager1;
            this.barDockControlBottom.Size = new System.Drawing.Size(572, 0);
            // 
            // barDockControlLeft
            // 
            this.barDockControlLeft.CausesValidation = false;
            this.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left;
            this.barDockControlLeft.Location = new System.Drawing.Point(0, 24);
            this.barDockControlLeft.Manager = this.barManager1;
            this.barDockControlLeft.Size = new System.Drawing.Size(0, 423);
            // 
            // barDockControlRight
            // 
            this.barDockControlRight.CausesValidation = false;
            this.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right;
            this.barDockControlRight.Location = new System.Drawing.Point(572, 24);
            this.barDockControlRight.Manager = this.barManager1;
            this.barDockControlRight.Size = new System.Drawing.Size(0, 423);
            // 
            // btn_timkiem
            // 
            this.btn_timkiem.Caption = "Tìm kiếm";
            this.btn_timkiem.Id = 3;
            this.btn_timkiem.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("btn_timkiem.ImageOptions.SvgImage")));
            this.btn_timkiem.Name = "btn_timkiem";
            // 
            // btn_thoat
            // 
            this.btn_thoat.Caption = "Thoát";
            this.btn_thoat.Id = 4;
            this.btn_thoat.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btn_thoat.ImageOptions.Image")));
            this.btn_thoat.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btn_thoat.ImageOptions.LargeImage")));
            this.btn_thoat.Name = "btn_thoat";
            // 
            // btn_export
            // 
            this.btn_export.Caption = "Export danh sách";
            this.btn_export.Id = 5;
            this.btn_export.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btn_export.ImageOptions.Image")));
            this.btn_export.Name = "btn_export";
            this.btn_export.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph;
            // 
            // btn_refresh
            // 
            this.btn_refresh.Caption = "Refresh";
            this.btn_refresh.Id = 6;
            this.btn_refresh.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btn_refresh.ImageOptions.Image")));
            this.btn_refresh.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btn_refresh.ImageOptions.LargeImage")));
            this.btn_refresh.Name = "btn_refresh";
            // 
            // barButtonItem1
            // 
            this.barButtonItem1.Id = 7;
            this.barButtonItem1.Name = "barButtonItem1";
            // 
            // btn_giayto
            // 
            this.btn_giayto.Caption = "Giấy tờ liên quan";
            this.btn_giayto.Id = 8;
            this.btn_giayto.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("btn_giayto.ImageOptions.SvgImage")));
            this.btn_giayto.Name = "btn_giayto";
            // 
            // barButtonItem2
            // 
            this.barButtonItem2.Caption = "Thoát";
            this.barButtonItem2.Id = 10;
            this.barButtonItem2.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("barButtonItem2.ImageOptions.Image")));
            this.barButtonItem2.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("barButtonItem2.ImageOptions.LargeImage")));
            this.barButtonItem2.Name = "barButtonItem2";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txt_magt);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.cbo_luutru);
            this.groupBox1.Controls.Add(this.btn__choose);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txt_url);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.txt_tengt);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(0, 24);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(572, 156);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Thông tin giấy tờ";
            // 
            // txt_magt
            // 
            this.txt_magt.Location = new System.Drawing.Point(151, 36);
            this.txt_magt.Name = "txt_magt";
            this.txt_magt.Size = new System.Drawing.Size(215, 21);
            this.txt_magt.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(84, 120);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Đối tượng";
            // 
            // cbo_luutru
            // 
            this.cbo_luutru.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbo_luutru.FormattingEnabled = true;
            this.cbo_luutru.Items.AddRange(new object[] {
            "Đề tài",
            "Hội thảo",
            "Cuộc thi",
            "Sinh viên",
            "Giảng viên"});
            this.cbo_luutru.Location = new System.Drawing.Point(151, 117);
            this.cbo_luutru.Name = "cbo_luutru";
            this.cbo_luutru.Size = new System.Drawing.Size(157, 21);
            this.cbo_luutru.TabIndex = 7;
            // 
            // btn__choose
            // 
            this.btn__choose.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.btn__choose.FlatAppearance.BorderSize = 0;
            this.btn__choose.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn__choose.Image = ((System.Drawing.Image)(resources.GetObject("btn__choose.Image")));
            this.btn__choose.Location = new System.Drawing.Point(420, 88);
            this.btn__choose.Name = "btn__choose";
            this.btn__choose.Size = new System.Drawing.Size(75, 23);
            this.btn__choose.TabIndex = 6;
            this.btn__choose.Text = "Choose";
            this.btn__choose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn__choose.UseVisualStyleBackColor = false;
            this.btn__choose.Click += new System.EventHandler(this.btn__choose_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(78, 92);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(61, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Đường dẫn";
            // 
            // txt_url
            // 
            this.txt_url.Location = new System.Drawing.Point(151, 89);
            this.txt_url.Name = "txt_url";
            this.txt_url.ReadOnly = true;
            this.txt_url.Size = new System.Drawing.Size(268, 21);
            this.txt_url.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(97, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(42, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Tiêu đề";
            // 
            // txt_tengt
            // 
            this.txt_tengt.Location = new System.Drawing.Point(151, 62);
            this.txt_tengt.Name = "txt_tengt";
            this.txt_tengt.Size = new System.Drawing.Size(343, 21);
            this.txt_tengt.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(82, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Mã giấy tờ";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dgv_giayto);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(0, 180);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(572, 267);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Danh sách giấy tờ";
            // 
            // dgv_giayto
            // 
            this.dgv_giayto.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_giayto.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_giayto.Location = new System.Drawing.Point(3, 17);
            this.dgv_giayto.Name = "dgv_giayto";
            this.dgv_giayto.ReadOnly = true;
            this.dgv_giayto.Size = new System.Drawing.Size(566, 247);
            this.dgv_giayto.TabIndex = 0;
            this.dgv_giayto.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_giayto_CellClick);
            // 
            // uc_giaytoluutru
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.barDockControlLeft);
            this.Controls.Add(this.barDockControlRight);
            this.Controls.Add(this.barDockControlBottom);
            this.Controls.Add(this.barDockControlTop);
            this.Name = "uc_giaytoluutru";
            this.Size = new System.Drawing.Size(572, 447);
            this.Load += new System.EventHandler(this.uc_giaytoluutru_Load);
            ((System.ComponentModel.ISupportInitialize)(this.barManager1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_giayto)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraBars.BarManager barManager1;
        private DevExpress.XtraBars.Bar bar2;
        private DevExpress.XtraBars.BarButtonItem btn_them;
        private DevExpress.XtraBars.BarButtonItem btn_sua;
        private DevExpress.XtraBars.BarButtonItem btn_xoa;
        private DevExpress.XtraBars.BarButtonItem btn_openfile;
        private DevExpress.XtraBars.BarDockControl barDockControlTop;
        private DevExpress.XtraBars.BarDockControl barDockControlBottom;
        private DevExpress.XtraBars.BarDockControl barDockControlLeft;
        private DevExpress.XtraBars.BarDockControl barDockControlRight;
        private DevExpress.XtraBars.BarButtonItem btn_timkiem;
        private DevExpress.XtraBars.BarButtonItem btn_thoat;
        private DevExpress.XtraBars.BarButtonItem btn_export;
        private DevExpress.XtraBars.BarButtonItem btn_refresh;
        private DevExpress.XtraBars.BarButtonItem barButtonItem1;
        private DevExpress.XtraBars.BarButtonItem btn_giayto;
        private DevExpress.XtraBars.BarButtonItem barButtonItem2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btn__choose;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_url;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_tengt;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView dgv_giayto;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cbo_luutru;
        private System.Windows.Forms.TextBox txt_magt;
    }
}
