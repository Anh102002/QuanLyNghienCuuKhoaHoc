
namespace QL_NCKH
{
    partial class QuenMatKhau
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(QuenMatKhau));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_user = new System.Windows.Forms.TextBox();
            this.txt_email = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_resetMk = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btn_khoiphuc = new System.Windows.Forms.Button();
            this.btn_thoat = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(23, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(254, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "KHÔI PHỤC MẬT KHẨU";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Username";
            // 
            // txt_user
            // 
            this.txt_user.Location = new System.Drawing.Point(31, 75);
            this.txt_user.Name = "txt_user";
            this.txt_user.Size = new System.Drawing.Size(236, 21);
            this.txt_user.TabIndex = 2;
            // 
            // txt_email
            // 
            this.txt_email.Location = new System.Drawing.Point(31, 123);
            this.txt_email.Name = "txt_email";
            this.txt_email.Size = new System.Drawing.Size(236, 21);
            this.txt_email.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(31, 103);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(31, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Email";
            // 
            // txt_resetMk
            // 
            this.txt_resetMk.BackColor = System.Drawing.Color.White;
            this.txt_resetMk.Location = new System.Drawing.Point(32, 171);
            this.txt_resetMk.Name = "txt_resetMk";
            this.txt_resetMk.ReadOnly = true;
            this.txt_resetMk.Size = new System.Drawing.Size(236, 21);
            this.txt_resetMk.TabIndex = 6;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(32, 151);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(102, 13);
            this.label4.TabIndex = 5;
            this.label4.Text = "Mật khẩu khôi phục ";
            // 
            // btn_khoiphuc
            // 
            this.btn_khoiphuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_khoiphuc.Image = ((System.Drawing.Image)(resources.GetObject("btn_khoiphuc.Image")));
            this.btn_khoiphuc.Location = new System.Drawing.Point(80, 216);
            this.btn_khoiphuc.Name = "btn_khoiphuc";
            this.btn_khoiphuc.Size = new System.Drawing.Size(95, 23);
            this.btn_khoiphuc.TabIndex = 7;
            this.btn_khoiphuc.Text = "Khôi phục";
            this.btn_khoiphuc.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_khoiphuc.UseVisualStyleBackColor = true;
            this.btn_khoiphuc.Click += new System.EventHandler(this.button1_Click);
            // 
            // btn_thoat
            // 
            this.btn_thoat.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_thoat.Image = ((System.Drawing.Image)(resources.GetObject("btn_thoat.Image")));
            this.btn_thoat.Location = new System.Drawing.Point(181, 216);
            this.btn_thoat.Name = "btn_thoat";
            this.btn_thoat.Size = new System.Drawing.Size(87, 23);
            this.btn_thoat.TabIndex = 8;
            this.btn_thoat.Text = "Thoát";
            this.btn_thoat.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btn_thoat.UseVisualStyleBackColor = true;
            this.btn_thoat.Click += new System.EventHandler(this.btn_thoat_Click);
            // 
            // QuenMatKhau
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(298, 271);
            this.Controls.Add(this.btn_thoat);
            this.Controls.Add(this.btn_khoiphuc);
            this.Controls.Add(this.txt_resetMk);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txt_email);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txt_user);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.IconOptions.Image = ((System.Drawing.Image)(resources.GetObject("QuenMatKhau.IconOptions.Image")));
            this.Name = "QuenMatKhau";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Quên mật khẩu";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_user;
        private System.Windows.Forms.TextBox txt_email;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_resetMk;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btn_khoiphuc;
        private System.Windows.Forms.Button btn_thoat;
    }
}