using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.SqlClient;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QL_NCKH
{
    public class XoaSV
    {
        MyClass my = new MyClass();
        public bool XoaSinhVien(string ma)
        {
            try
            {
                string sql = "delete from SinhVien where MaSV = '" + ma + "' ";
                int up = my.Update(sql);
                if (up > 0)
                {
                    return true;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Lỗi xóa sinh viên {"+ex.Message+"}","Lỗi");
            }

            
            return false;
        }

        public bool XoaGTSV(string ma)
        {          
            try
            {
                string sql = "delete from GiayToSV where MaSV = @Masv ";
                SqlCommand command = my.SqlCommand(sql);
                command.Parameters.AddWithValue("@Masv", ma);
                int up = command.ExecuteNonQuery();
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa giấy tờ sinh viên {" + ex.Message + "}", "Lỗi");
            }
            return false;
        }

        public bool XoaSVDT(string ma)
        {
            try
            {
                string sql = "delete from ChiTietSVDeTai where MaSV = @Masv ";
                SqlCommand command = my.SqlCommand(sql);
                command.Parameters.AddWithValue("@Masv", ma);
                int up = command.ExecuteNonQuery();
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa sinh viên thamg gia đề tài {" + ex.Message + "}", "Lỗi");
            }
            return false;
        }

        public bool XoaTVCT(string ma)
        {
            try
            {
                string sql = "delete from ThanhVienCuocThi where MaSV = @Masv ";
                SqlCommand command = my.SqlCommand(sql);
                command.Parameters.AddWithValue("@Masv", ma);
                int up = command.ExecuteNonQuery();
                if (up >= 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi xóa sinh viên thamg gia đề tài {" + ex.Message + "}", "Lỗi");
            }
            return false;
        }
    }
}
