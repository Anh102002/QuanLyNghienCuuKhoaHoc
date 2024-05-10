using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraBars;

namespace QL_NCKH
{
    public partial class MainFrm : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        MyClass my = new MyClass();
        private string quyen;
        private string Hoten;
        private string username;
        private int Id;

        uc_giangvien ucgv;
        uc_Account uc_Account;
        uc_phiendangnhap uc_Phiendangnhap;
        uc_sinhvien uc_sv;
        uc_detaiquocgia ucdtqg;
        uc_tvngoaitruong tvnt;
        uc_detaicapbo dtcb;
        uc_detaicapCoSo dtcs;
        uc_detaiNCKHsv dtsv;
        uc_capnhatketqua capnhat;
        uc_tapchiNCKH tapchi;
        uc_baibaoquocte bbquote;
        uc_baibaotrongnuoc bbtrongnnuoc;
        uc_cuocthisangtao ucsangtao;
        uc_cuocthiSTKNcaptruong ucstknsaptruong;
        uc_capnhatcuocthi uckqct;
        uc_hoithaoquocte uchtqt;
        uc_hoithaoquocgia uchtqg;
        uc_hoithaocapbo uchtcb;
        uc_hoithaocaptruong uchtct;
        uc_hoithaocapkhoa uchtck;
        uc_sinhvienngoaitruong ucsvnt;
        uc_giaytoluutru ucgtlt;
        uc_thongke uc_thongke;

        public string getQuyen()
        {
            return this.quyen;
        }
        public string getHoten()
        {
            return this.Hoten;
        }
        public string getUser()
        {
            return this.username;
        }
        public int getId()
        {
            return this.Id;
        }

        public void setHoten(string hoten)
        {
            this.Hoten = hoten;
        }
        public void setQuyen(string quyen)
        {
            this.quyen = quyen;
        }
        public void setUser(string user)
        {
            this.username = user;
        }

        public void setId(int id)
        {
            this.Id = id;
        }

        bool isThoat = true;

        public MainFrm()
        {
            InitializeComponent();
        }

        private void barButtonItem2_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (isThoat)
            {
                Application.Exit();

            }
        }

        private void barButtonItem3_ItemClick(object sender, ItemClickEventArgs e)
        {
            uc_DoiMatKhau f = new uc_DoiMatKhau();
            string user = getUser();
            f.setUser(user);
            f.ShowDialog();
        }
        public void logout()
        {
            try
            {
                int id = getId();
                string time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string sql = "update PhienDangNhap set ThoiGianDangXuat= '" + time + "' where Id = '" + id + "'   ";
                my.Update(sql);
            }
            catch
            {
                MessageBox.Show("Lỗi kiểm tra đăng xuất", "Thông báo", MessageBoxButtons.OKCancel);

            }
        }
        private void barButtonItem4_ItemClick(object sender, ItemClickEventArgs e)
        {
            isThoat = false;
            this.Close();
            logout();
            this.Hide();
            Login fr = new Login();
            fr.ShowDialog();
        }

        private void barButtonItem6_ItemClick(object sender, ItemClickEventArgs e)
        {
            string check = "ttgv";


            DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, check);

            if (existingTabPage != null)
            {
                xtraTabControl1.SelectedTabPage = existingTabPage;
            }
            else
            {

                ucgv = new uc_giangvien();


                DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                tabPage.Text = btn_giangvien.Caption;

                ucgv.Dock = DockStyle.Fill;
                tabPage.Controls.Add(ucgv);

                tabPage.Tag = check;
                xtraTabControl1.TabPages.Add(tabPage);


                xtraTabControl1.SelectedTabPage = tabPage;
            }
        }

        private void barButtonItem7_ItemClick(object sender, ItemClickEventArgs e)
        {
            string checksv = "sv";


            DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checksv);

            if (existingTabPage != null)
            {
                xtraTabControl1.SelectedTabPage = existingTabPage;
            }
            else
            {

                uc_sv = new uc_sinhvien();


                DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                tabPage.Text = btn_sinhvien.Caption;

                uc_sv.Dock = DockStyle.Fill;
                tabPage.Controls.Add(uc_sv);

                tabPage.Tag = checksv;
                xtraTabControl1.TabPages.Add(tabPage);


                xtraTabControl1.SelectedTabPage = tabPage;
            }
        }

        private void barButtonItem5_ItemClick(object sender, ItemClickEventArgs e)
        {
            string maDeTaiToCheck = "YourUniqueIdentifier";


            DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, maDeTaiToCheck);

            if (existingTabPage != null)
            {
                xtraTabControl1.SelectedTabPage = existingTabPage;
            }
            else
            {

                ucdtqg = new uc_detaiquocgia();


                DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                tabPage.Text = btn_detaiquocgia.Caption;

                ucdtqg.Dock = DockStyle.Fill;
                tabPage.Controls.Add(ucdtqg);

                tabPage.Tag = maDeTaiToCheck;
                xtraTabControl1.TabPages.Add(tabPage);


                xtraTabControl1.SelectedTabPage = tabPage;
            }
        }

        private void barButtonItem8_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkdtcb = "dtcb";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkdtcb);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {

                    dtcb = new uc_detaicapbo();


                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = btn_detaicapbo.Caption;

                    dtcb.Dock = DockStyle.Left;
                    tabPage.Controls.Add(dtcb);

                    tabPage.Tag = checkdtcb;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở danh sách đề tài cấp bộ ", "Lỗi");
            }
        }

        private void barCheckItem1_CheckedChanged(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkdtcs = "dtcs";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkdtcs);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {


                    dtcs = new uc_detaicapCoSo();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = btn_detaicapcoso.Caption;

                    dtcs.Dock = DockStyle.Left;
                    tabPage.Controls.Add(dtcs);

                    tabPage.Tag = checkdtcs;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở danh sách đề tài cấp cơ sở ", "Lỗi");
            }
        }

        private void barButtonItem9_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkdtsv = "dtsv";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkdtsv);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {


                    dtsv = new uc_detaiNCKHsv();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = btn_detaiNCKHsv.Caption;

                    dtsv.Dock = DockStyle.Left;
                    tabPage.Controls.Add(dtsv);

                    tabPage.Tag = checkdtsv;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở danh sách đề tài NCKH sinh viên ", "Lỗi");
            }
        }

        private void barButtonItem10_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkcapnhat = "cndt";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkcapnhat);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {

                    capnhat = new uc_capnhatketqua();


                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = btn_capnhatdetai.Caption;

                    capnhat.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(capnhat);

                    tabPage.Tag = checkcapnhat;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở danh sách kết quả nghiệm thu ", "Lỗi");
            }
        }

        private void barButtonItem11_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkctst = "ctst";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkctst);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {

                    ucsangtao = new uc_cuocthisangtao();


                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = btn_cuocthisangtao.Caption;

                    ucsangtao.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(ucsangtao);

                    tabPage.Tag = checkctst;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở cuộc thi sáng tạo ", "Lỗi");
            }
        }
        private DevExpress.XtraTab.XtraTabPage FindTabPageByTag(DevExpress.XtraTab.XtraTabControl tabControl, object tag)
        {
            foreach (DevExpress.XtraTab.XtraTabPage tabPage in tabControl.TabPages)
            {
                if (tabPage.Tag != null && tabPage.Tag.Equals(tag))
                {
                    return tabPage;
                }
            }
            return null;
        }
        private void barButtonItem13_ItemClick(object sender, ItemClickEventArgs e)
        {
            string checkaccount = "account";


            DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkaccount);

            if (existingTabPage != null)
            {
                xtraTabControl1.SelectedTabPage = existingTabPage;
            }
            else
            {

                uc_Account = new uc_Account();


                DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                tabPage.Text = btn_taikhoan.Caption;

                uc_Account.Dock = DockStyle.Fill;
                tabPage.Controls.Add(uc_Account);

                tabPage.Tag = checkaccount;
                xtraTabControl1.TabPages.Add(tabPage);


                xtraTabControl1.SelectedTabPage = tabPage;
            }
        }

        private void barButtonItem15_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkquocte = "bbqt";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkquocte);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {


                    bbquote = new uc_baibaoquocte();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = btn_baibaoquocte.Caption;

                    bbquote.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(bbquote);

                    tabPage.Tag = checkquocte;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở danh sách bài báo quốc tế ", "Lỗi");
            }
        }

        private void btn_lsdangnhap_ItemClick(object sender, ItemClickEventArgs e)
        {
            string checkdn = "login";


            DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkdn);

            if (existingTabPage != null)
            {
                xtraTabControl1.SelectedTabPage = existingTabPage;
            }
            else
            {

                uc_Phiendangnhap = new uc_phiendangnhap();


                DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                tabPage.Text = btn_lsdangnhap.Caption;

                uc_Phiendangnhap.Dock = DockStyle.Fill;
                tabPage.Controls.Add(uc_Phiendangnhap);

                tabPage.Tag = checkdn;
                xtraTabControl1.TabPages.Add(tabPage);


                xtraTabControl1.SelectedTabPage = tabPage;
            }
        }

        private void xtraTabControl1_CloseButtonClick(object sender, EventArgs e)
        {
            if (xtraTabControl1.SelectedTabPage != null)
            {

                DevExpress.XtraTab.XtraTabPage selectedTabPage = xtraTabControl1.SelectedTabPage;
                xtraTabControl1.TabPages.Remove(selectedTabPage);
            }
        }

        private void MainFrm_Load(object sender, EventArgs e)
        {
            if (getQuyen() == "Administrators")
            {
                Page8.Visible = true;
                

            }
            else
            if (getQuyen() == "User")
            {

                Page8.Visible = false;
            }

            txt_tentk.Caption = getHoten();

           
           
            
        }

        private void btn_baibaotrongnuoc_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checktrongnuoc = "bbTN";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checktrongnuoc);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {

                    bbtrongnnuoc = new uc_baibaotrongnuoc();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = btn_baibaotrongnuoc.Caption;

                    bbtrongnnuoc.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(bbtrongnnuoc);

                    tabPage.Tag = checktrongnuoc;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở danh sách bài báo trong nước", "Lỗi");
            }
        }

        private void btn_tapchiNCKH_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checktapchi = "cndt";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checktapchi);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {


                    tapchi = new uc_tapchiNCKH();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = btn_tapchiNCKH.Caption;

                    tapchi.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(tapchi);

                    tabPage.Tag = checktapchi;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở danh sách tạp chí NCKH ", "Lỗi");
            }
        }

        private void MainFrm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (isThoat)
            {                
                logout();                    
            }
        }

        private void MainFrm_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Dispose();
        }

        private void barStaticItem4_ItemClick(object sender, ItemClickEventArgs e)
        {

        }

        private void btn_cuocthikhoinghiep_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkSTKNCT = "stknct";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkSTKNCT);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {


                    ucstknsaptruong = new uc_cuocthiSTKNcaptruong();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = btn_cuocthikhoinghiep.Caption;

                    ucstknsaptruong.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(ucstknsaptruong);

                    tabPage.Tag = checkSTKNCT;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở cuộc thi sáng tạo khởi nghiệp cấp trường ", "Lỗi");
            }
        }

        private void xtraTabControl1_Click(object sender, EventArgs e)
        {

        }

        private void barButtonItem3_ItemClick_1(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkKQCT = "kqct";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkKQCT);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {


                    uckqct = new uc_capnhatcuocthi();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = barButtonItem3.Caption;

                    uckqct.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(uckqct);

                    tabPage.Tag = checkKQCT;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở cuộc thi sáng tạo khởi nghiệp cấp trường ", "Lỗi");
            }
        }

        private void barButtonItem4_ItemClick_1(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkCBNT = "cbnt";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkCBNT);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {


                    tvnt = new uc_tvngoaitruong();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = barButtonItem4.Caption;

                    tvnt.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(tvnt);

                    tabPage.Tag = checkCBNT;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở cán bộ ngoài trường ", "Lỗi");
            }
        }

        private void barButtonItem7_ItemClick_1(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkHTQT = "htqt";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkHTQT);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {


                    uchtqt = new uc_hoithaoquocte();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = barButtonItem7.Caption;

                    uchtqt.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(uchtqt);

                    tabPage.Tag = checkHTQT;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở hội thảo quốc tế ", "Lỗi");
            }
        }

        private void barButtonItem8_ItemClick_1(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkHTQG = "htqg";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkHTQG);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {


                    uchtqg = new uc_hoithaoquocgia();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = barButtonItem8.Caption;

                    uchtqg.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(uchtqg);

                    tabPage.Tag = checkHTQG;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở hội thảo cấp nhà nước ", "Lỗi");
            }
        }

        private void barButtonItem9_ItemClick_1(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkHTCB = "htcb";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkHTCB);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {


                    uchtcb = new uc_hoithaocapbo();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = barButtonItem9.Caption;

                    uchtcb.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(uchtcb);

                    tabPage.Tag = checkHTCB;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở hội thảo cấp bộ ", "Lỗi");
            }
        }

        private void barButtonItem10_ItemClick_1(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkHTCT = "htct";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkHTCT);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {

                    uchtct = new uc_hoithaocaptruong();



                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = barButtonItem10.Caption;

                    uchtct.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(uchtct);

                    tabPage.Tag = checkHTCT;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở hội thảo cấp trường ", "Lỗi");
            }
        }

        private void barButtonItem11_ItemClick_1(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkHTCK = "htck";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkHTCK);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {



                    uchtck = new uc_hoithaocapkhoa();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = barButtonItem11.Caption;

                    uchtck.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(uchtck);

                    tabPage.Tag = checkHTCK;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở hội thảo cấp khoa ", "Lỗi");
            }
        }

        private void barButtonItem6_ItemClick_1(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkSVNT = "svnt";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkSVNT);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {



                    ucsvnt = new uc_sinhvienngoaitruong();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = barButtonItem6.Caption;

                    ucsvnt.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(ucsvnt);

                    tabPage.Tag = checkSVNT;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở sinh viên ngoài trường ", "Lỗi");
            }
        }

        private void barButtonItem12_ItemClick(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkGTLT = "gtlt";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkGTLT);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {



                    ucgtlt = new uc_giaytoluutru();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = barButtonItem12.Caption;

                    ucgtlt.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(ucgtlt);

                    tabPage.Tag = checkGTLT;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở giấy tờ lưu trữ ", "Lỗi");
            }
        }

        private void barButtonItem13_ItemClick_1(object sender, ItemClickEventArgs e)
        {
            try
            {
                string checkTK = "thongke";


                DevExpress.XtraTab.XtraTabPage existingTabPage = FindTabPageByTag(xtraTabControl1, checkTK);

                if (existingTabPage != null)
                {
                    xtraTabControl1.SelectedTabPage = existingTabPage;
                }
                else
                {



                    uc_thongke = new uc_thongke();

                    DevExpress.XtraTab.XtraTabPage tabPage = new DevExpress.XtraTab.XtraTabPage();
                    tabPage.Text = barButtonItem13.Caption;

                    uc_thongke.Dock = DockStyle.Fill;
                    tabPage.Controls.Add(uc_thongke);

                    tabPage.Tag = checkTK;
                    xtraTabControl1.TabPages.Add(tabPage);


                    xtraTabControl1.SelectedTabPage = tabPage;
                }
            }
            catch
            {
                MessageBox.Show("$ Lỗi mở thống kê nghiên cứu khoa học ", "Lỗi");
            }
        }
    }
}