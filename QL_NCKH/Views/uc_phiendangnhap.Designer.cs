
namespace QL_NCKH
{
    partial class uc_phiendangnhap
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_phiendangnhap));
            this.barManager1 = new DevExpress.XtraBars.BarManager(this.components);
            this.bar2 = new DevExpress.XtraBars.Bar();
            this.btn_xoa = new DevExpress.XtraBars.BarButtonItem();
            this.btn_refresh = new DevExpress.XtraBars.BarButtonItem();
            this.barDockControlTop = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlBottom = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlLeft = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlRight = new DevExpress.XtraBars.BarDockControl();
            this.btn_them = new DevExpress.XtraBars.BarButtonItem();
            this.btn_sua = new DevExpress.XtraBars.BarButtonItem();
            this.btn_timkiem = new DevExpress.XtraBars.BarButtonItem();
            this.btn_thoat = new DevExpress.XtraBars.BarButtonItem();
            this.btn_export = new DevExpress.XtraBars.BarButtonItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.txt_timkiem = new DevExpress.XtraEditors.SearchControl();
            this.label5 = new System.Windows.Forms.Label();
            this.cb_loai = new System.Windows.Forms.ComboBox();
            this.dtp_logout = new System.Windows.Forms.DateTimePicker();
            this.dtp_login = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_user = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_id = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dgv_dangnhap = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.barManager1)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txt_timkiem.Properties)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_dangnhap)).BeginInit();
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
            this.btn_refresh});
            this.barManager1.MaxItemId = 7;
            // 
            // bar2
            // 
            this.bar2.BarName = "Tools";
            this.bar2.DockCol = 0;
            this.bar2.DockRow = 0;
            this.bar2.DockStyle = DevExpress.XtraBars.BarDockStyle.Top;
            this.bar2.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
            new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.btn_xoa, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph),
            new DevExpress.XtraBars.LinkPersistInfo(DevExpress.XtraBars.BarLinkUserDefines.PaintStyle, this.btn_refresh, "", true, true, true, 0, null, DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph)});
            this.bar2.Text = "Tools";
            // 
            // btn_xoa
            // 
            this.btn_xoa.Caption = "Xóa";
            this.btn_xoa.Id = 2;
            this.btn_xoa.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("btn_xoa.ImageOptions.SvgImage")));
            this.btn_xoa.Name = "btn_xoa";
            this.btn_xoa.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btn_xoa_ItemClick);
            // 
            // btn_refresh
            // 
            this.btn_refresh.Caption = "Refresh";
            this.btn_refresh.Id = 6;
            this.btn_refresh.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btn_refresh.ImageOptions.Image")));
            this.btn_refresh.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btn_refresh.ImageOptions.LargeImage")));
            this.btn_refresh.Name = "btn_refresh";
            this.btn_refresh.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btn_refresh_ItemClick);
            // 
            // barDockControlTop
            // 
            this.barDockControlTop.Appearance.Font = new System.Drawing.Font("Century Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.barDockControlTop.Appearance.Options.UseFont = true;
            this.barDockControlTop.CausesValidation = false;
            this.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.barDockControlTop.Location = new System.Drawing.Point(0, 0);
            this.barDockControlTop.Manager = this.barManager1;
            this.barDockControlTop.Size = new System.Drawing.Size(755, 39);
            // 
            // barDockControlBottom
            // 
            this.barDockControlBottom.CausesValidation = false;
            this.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.barDockControlBottom.Location = new System.Drawing.Point(0, 461);
            this.barDockControlBottom.Manager = this.barManager1;
            this.barDockControlBottom.Size = new System.Drawing.Size(755, 0);
            // 
            // barDockControlLeft
            // 
            this.barDockControlLeft.CausesValidation = false;
            this.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left;
            this.barDockControlLeft.Location = new System.Drawing.Point(0, 39);
            this.barDockControlLeft.Manager = this.barManager1;
            this.barDockControlLeft.Size = new System.Drawing.Size(0, 422);
            // 
            // barDockControlRight
            // 
            this.barDockControlRight.CausesValidation = false;
            this.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right;
            this.barDockControlRight.Location = new System.Drawing.Point(755, 39);
            this.barDockControlRight.Manager = this.barManager1;
            this.barDockControlRight.Size = new System.Drawing.Size(0, 422);
            // 
            // btn_them
            // 
            this.btn_them.Caption = "Thêm\r\n";
            this.btn_them.Id = 0;
            this.btn_them.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btn_them.ImageOptions.Image")));
            this.btn_them.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("btn_them.ImageOptions.LargeImage")));
            this.btn_them.Name = "btn_them";
            // 
            // btn_sua
            // 
            this.btn_sua.Caption = "Sửa";
            this.btn_sua.Id = 1;
            this.btn_sua.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("btn_sua.ImageOptions.SvgImage")));
            this.btn_sua.Name = "btn_sua";
            this.btn_sua.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.btn_sua_ItemClick);
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
            // btn_export
            // 
            this.btn_export.Caption = "Export danh sách";
            this.btn_export.Id = 5;
            this.btn_export.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("btn_export.ImageOptions.Image")));
            this.btn_export.Name = "btn_export";
            this.btn_export.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.txt_timkiem);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.cb_loai);
            this.groupBox1.Controls.Add(this.dtp_logout);
            this.groupBox1.Controls.Add(this.dtp_login);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txt_user);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.txt_id);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(4, 40);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(749, 169);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Thông tin đăng nhập";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.button1.Image = global::QL_NCKH.Properties.Resources.search_16;
            this.button1.Location = new System.Drawing.Point(671, 136);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 93;
            this.button1.Text = "Tìm kiếm";
            this.button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txt_timkiem
            // 
            this.txt_timkiem.Location = new System.Drawing.Point(270, 133);
            this.txt_timkiem.MenuManager = this.barManager1;
            this.txt_timkiem.Name = "txt_timkiem";
            this.txt_timkiem.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Repository.ClearButton()});
            this.txt_timkiem.Properties.ShowSearchButton = false;
            this.txt_timkiem.Size = new System.Drawing.Size(395, 28);
            this.txt_timkiem.TabIndex = 92;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 140);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(73, 13);
            this.label5.TabIndex = 12;
            this.label5.Text = "Tìm kiếm theo";
            // 
            // cb_loai
            // 
            this.cb_loai.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_loai.FormattingEnabled = true;
            this.cb_loai.Items.AddRange(new object[] {
            "ID",
            "Username"});
            this.cb_loai.Location = new System.Drawing.Point(89, 137);
            this.cb_loai.Name = "cb_loai";
            this.cb_loai.Size = new System.Drawing.Size(175, 21);
            this.cb_loai.TabIndex = 10;
            // 
            // dtp_logout
            // 
            this.dtp_logout.CustomFormat = "dd\\MM\\yyyy HH:mm:ss";
            this.dtp_logout.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtp_logout.Location = new System.Drawing.Point(330, 91);
            this.dtp_logout.Name = "dtp_logout";
            this.dtp_logout.Size = new System.Drawing.Size(200, 20);
            this.dtp_logout.TabIndex = 9;
            // 
            // dtp_login
            // 
            this.dtp_login.CustomFormat = "dd\\MM\\yyyy HH:mm:ss";
            this.dtp_login.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtp_login.Location = new System.Drawing.Point(330, 50);
            this.dtp_login.Name = "dtp_login";
            this.dtp_login.Size = new System.Drawing.Size(200, 20);
            this.dtp_login.TabIndex = 8;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(327, 74);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(102, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Thời gian đăng xuất";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(327, 33);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(106, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Thời gian đăng nhập";
            // 
            // txt_user
            // 
            this.txt_user.BackColor = System.Drawing.Color.White;
            this.txt_user.Location = new System.Drawing.Point(89, 91);
            this.txt_user.Name = "txt_user";
            this.txt_user.ReadOnly = true;
            this.txt_user.Size = new System.Drawing.Size(175, 20);
            this.txt_user.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(86, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Username";
            // 
            // txt_id
            // 
            this.txt_id.BackColor = System.Drawing.Color.White;
            this.txt_id.Location = new System.Drawing.Point(89, 50);
            this.txt_id.Name = "txt_id";
            this.txt_id.ReadOnly = true;
            this.txt_id.Size = new System.Drawing.Size(175, 20);
            this.txt_id.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(86, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(16, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Id";
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.dgv_dangnhap);
            this.groupBox2.Location = new System.Drawing.Point(4, 226);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(748, 232);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Danh sách đăng nhập";
            // 
            // dgv_dangnhap
            // 
            this.dgv_dangnhap.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_dangnhap.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_dangnhap.Location = new System.Drawing.Point(3, 16);
            this.dgv_dangnhap.Name = "dgv_dangnhap";
            this.dgv_dangnhap.ReadOnly = true;
            this.dgv_dangnhap.Size = new System.Drawing.Size(742, 213);
            this.dgv_dangnhap.TabIndex = 0;
            this.dgv_dangnhap.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_dangnhap_CellClick);
            this.dgv_dangnhap.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_dangnhap_CellContentClick);
            // 
            // uc_phiendangnhap
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.barDockControlLeft);
            this.Controls.Add(this.barDockControlRight);
            this.Controls.Add(this.barDockControlBottom);
            this.Controls.Add(this.barDockControlTop);
            this.Name = "uc_phiendangnhap";
            this.Size = new System.Drawing.Size(755, 461);
            this.Load += new System.EventHandler(this.uc_phiendangnhap_Load);
            ((System.ComponentModel.ISupportInitialize)(this.barManager1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txt_timkiem.Properties)).EndInit();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_dangnhap)).EndInit();
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
        private DevExpress.XtraBars.BarDockControl barDockControlTop;
        private DevExpress.XtraBars.BarDockControl barDockControlBottom;
        private DevExpress.XtraBars.BarDockControl barDockControlLeft;
        private DevExpress.XtraBars.BarDockControl barDockControlRight;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView dgv_dangnhap;
        private System.Windows.Forms.GroupBox groupBox1;
        private DevExpress.XtraBars.BarButtonItem btn_thoat;
        private DevExpress.XtraBars.BarButtonItem btn_export;
        private System.Windows.Forms.DateTimePicker dtp_logout;
        private System.Windows.Forms.DateTimePicker dtp_login;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_user;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_id;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cb_loai;
        private DevExpress.XtraBars.BarButtonItem btn_refresh;
        private System.Windows.Forms.Button button1;
        private DevExpress.XtraEditors.SearchControl txt_timkiem;
    }
}
