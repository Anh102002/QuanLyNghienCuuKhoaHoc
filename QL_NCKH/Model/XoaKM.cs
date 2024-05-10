using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QL_NCKH
{
    class XoaKM
    {
        
        MyClass my = new MyClass();
        public bool XoaHDNT(string ma)
        {
            try
            {
                string sql = "delete from HoiDongNgoaiTruong where MaHD = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa hội đồng ngoài trường {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }

        public bool XoaBGKKM(string ma)
        {
            try
            {
                string sql = "delete from BGKCuocThiKM where MaKM = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa BGK ngoài trường {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }

        public bool XoaBLDNT(string ma)
        {
            try
            {
                string sql = "delete from BanLanhDaoCTNT where MaKM = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa BLD ngoài trường {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }

        public bool XoaTDHTKM(string ma)
        {
            try
            {
                string sql = "delete from TDHT_KhachMoi where MaKM = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa tham gia hội thảo phía khách mời {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }


        public bool XoaTCHTKM(string ma)
        {
            try
            {
                string sql = "delete from BanToChucHTNT where MaKM = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa BTC khách mời {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }

        public bool XoaKhachMoi(string ma)
        {
            try
            {
                string sql = "delete from TVNgoaiTruong where MaKM = '" + ma + "' ";
                int up = my.Update(sql);
                if (up > 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa khách mời {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }


        public bool XoaHD(string ma)
        {
            try
            {
                string sql = "delete from HopDong where MaKM = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa hợp đồng {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }

    }
}
