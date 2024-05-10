using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QL_NCKH.Model
{
    public class XoaHT
    {

        
        MyClass my = new MyClass();
        public bool XoaBTC(string ma)
        {
            try
            {
                string sql = "delete from BanToChucHT where MaHT=@Ma ";
                SqlCommand command = my.SqlCommand(sql);
                command.Parameters.AddWithValue("@Ma", ma);

                int up = command.ExecuteNonQuery();
                if (up >= 0)
                {
                    return true;
                }


            }
            catch(Exception ex)
            {
                MessageBox.Show("Lỗi ! xóa ban tổ chức {1}", "Lỗi");
            }

            return false;

        }
        public bool XoaBTCNT(string ma)
        {
            try
            {
                string sql = "delete from BanToChucHTNT  where MaHT=@Ma ";
                SqlCommand command = my.SqlCommand(sql);
                command.Parameters.AddWithValue("@Ma", ma);


                int up = command.ExecuteNonQuery();
                if (up >= 0)
                {
                    return true;
                }


            }
            catch
            {
                MessageBox.Show("Lỗi ! xóa ban tổ chức {2}", "Lỗi");
            }

            return false;

        }

        public bool XoaBBHT(string ma)
        {
            try
            {
                string sql = "delete from BaibaoHT  where MaHT=@Ma ";
                SqlCommand command = my.SqlCommand(sql);
                command.Parameters.AddWithValue("@Ma", ma);


                int up = command.ExecuteNonQuery();
                if (up >= 0)
                {
                    return true;
                }


            }
            catch
            {
                MessageBox.Show("Lỗi ! xóa bài báo hộ thảo", "Lỗi");
            }

            return false;

        }

        public bool XoaNT(string ma)
        {
            try
            {
                string sql = "delete from TDHT_PhiaNhaTruong  where MaHT=@Ma ";
                SqlCommand command = my.SqlCommand(sql);
                command.Parameters.AddWithValue("@Ma", ma);


                int up = command.ExecuteNonQuery();
                if (up >= 0)
                {
                    return true;
                }


            }
            catch
            {
                MessageBox.Show("Lỗi ! xóa phía nhà trường", "Lỗi");
            }

            return false;

        }
        public bool XoaKM(string ma)
        {
            try
            {
                string sql = "delete from TDHT_KhachMoi  where MaHT=@Ma ";
                SqlCommand command = my.SqlCommand(sql);
                command.Parameters.AddWithValue("@Ma", ma);


                int up = command.ExecuteNonQuery();
                if (up >= 0)
                {
                    return true;
                }


            }
            catch
            {
                MessageBox.Show("Lỗi ! xóa phía khách mời", "Lỗi");
            }

            return false;

        }
        public bool XoaDD(string ma)
        {
            try
            {
                string sql = "delete from TDHT_PhiaDaiDien where MaHT=@Ma ";
                SqlCommand command = my.SqlCommand(sql);
                command.Parameters.AddWithValue("@Ma", ma);


                int up = command.ExecuteNonQuery();
                if (up >= 0)
                {
                    return true;
                }


            }
            catch
            {
                MessageBox.Show("Lỗi ! xóa phía đại diện", "Lỗi");
            }

            return false;

        }
        public bool XoaCG(string ma)
        {
            try
            {
                string sql = "delete from TDHT_ChuyenGia  where MaHT=@Ma ";
                SqlCommand command = my.SqlCommand(sql);
                command.Parameters.AddWithValue("@Ma", ma);


                int up = command.ExecuteNonQuery();
                if (up >= 0)
                {
                    return true;
                }


            }
            catch
            {
                MessageBox.Show("Lỗi ! xóa phía chuyên gia", "Lỗi");
            }

            return false;

        }

        public bool XoaHoiThao(string ma)
        {
            try
            {
                string sql = "delete from HoiThao where MaHT=@Ma ";
                SqlCommand command = my.SqlCommand(sql);
                command.Parameters.AddWithValue("@Ma", ma);


                int up = command.ExecuteNonQuery();
                if (up > 0)
                {
                    return true;
                }


            }
            catch(Exception ex)
            {
                MessageBox.Show("Lỗi ! xóa hội thảo {"+ex.Message+"}", "Lỗi");
            }

            return false;

        }

    }
}
