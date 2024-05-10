using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QL_NCKH
{
    class XoaGV
    {
        MyClass my = new MyClass();
        public bool XoaGiangVien(string ma)
        {
            try
            {
                string sql = "delete from GiangVien where MaGV = '" + ma + "' ";
                int up = my.Update(sql);
                if (up > 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa giảng viên {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }


        public bool XoaGTGV(string ma)
        {
            try
            {
                string sql = "delete from GiayToGV where MaGV = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >=0 )
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa giấy tờ giảng viên {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }


        public bool XoaCTTC(string ma)
        {
            try
            {
                string sql = "delete from CTTapChi where MaGV = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa tác giả tạp chí giảng viên {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }

        public bool XoaGVHDCT(string ma)
        {
            try
            {
                string sql = "delete from GVHDCuocThi where MaGV = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa GVHD {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }


        public bool XoaBLDCT(string ma)
        {
            try
            {
                string sql = "delete from BanLanhDaoCT where MaGV = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa BLD cuộc thi {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }


        public bool XoaBGKCT(string ma)
        {
            try
            {
                string sql = "delete from BGKCuocThi where MaGV = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa BGK cuộc thi {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }

        public bool XoaHD(string ma)
        {
            try
            {
                string sql = "delete from HoiDong where MaGV = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa hội đồng {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }

        public bool XoaGVDT(string ma)
        {
            try
            {
                string sql = "delete from ChiTietGVDeTai where MaGV = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa giảng viên tham gia đề tài {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }


        public bool XoaTGBB(string ma)
        {
            try
            {
                string sql = "delete from TacGiaBaiBao where MaGV = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa tác giả bài báo {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }



        public bool XoaBTCCT(string ma)
        {
            try
            {
                string sql = "delete from BanToChucCT where MaGV = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa BTC cuộc thi {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }

        public bool XoaBHTCT(string ma)
        {
            try
            {
                string sql = "delete from BanHoTroCT where MaGV = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa BHTKT cuộc thi {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }


        public bool XoaBTCHT(string ma)
        {
            try
            {
                string sql = "delete from BanToChucHT where MaGV = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa BTC hội thảo {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }

        public bool XoaTDHTNT(string ma)
        {
            try
            {
                string sql = "delete from TDHT_PhiaNhaTruong where MaGV = '" + ma + "' ";
                int up = my.Update(sql);
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa TDHT về phía nhà trường {" + ex.Message + "}", "Lỗi");
            }


            return false;
        }
    }
}
