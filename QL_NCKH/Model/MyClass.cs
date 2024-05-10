using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Configuration;
using System.IO;
using Newtonsoft.Json;

namespace QL_NCKH
{
    public class MyClass
    {



        static SqlConnection sqlcon;
        string Connection = Properties.Settings.Default.qlnckh;
        //string Connection = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;
        //string Connection = ConfigurationManager.ConnectionStrings["MyDatabase"].ConnectionString;


        //public string GetConnectionString()
        //{
        //    try
        //    {
        //        string json = File.ReadAllText("appsettings.json");
        //        dynamic jsonObj = JsonConvert.DeserializeObject(json);
        //        return jsonObj.ConnectionStrings.MyConnectionString;
        //    }
        //    catch(Exception ex)
        //    {
        //        //System.Diagnostics.Debug.WriteLine("Error: " + ex.Message);
        //        MessageBox.Show("Lỗi kết nối .Vui lòng thử lại sau ", "Cảnh báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
        //    }
        //    return null;
        //}




        public void ketNoi()
        {
            //sqlcon.ConnectionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=QuanLyNCKH;Integrated Security=True;User ID= sa ;Password = 1234";

            try
            {


                sqlcon = new SqlConnection(Connection);

                sqlcon.Open();


            }
            catch
            {
                MessageBox.Show("Lỗi kết nối .Vui lòng thử lại sau ", "Cảnh báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                return;
            }

        }




        public SqlConnection ketNoiOnce()
        {

            try
            {               
                sqlcon = new SqlConnection(Connection);

                sqlcon.Open();

                return sqlcon;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error: " + ex.Message);
                MessageBox.Show("Lỗi kết nối .Vui lòng thử lại sau ", "Cảnh báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);



            }

            return null;

        }

        public void dongKetNoi()
        {
            if(sqlcon.State == System.Data.ConnectionState.Open)
            {
                sqlcon.Close();
                sqlcon.Dispose();
                sqlcon = null;
            }    
        }
        
        public DataTable DocDL(string sql)
        {
            //string Connection = GetConnectionString();
            //sqlcon = new SqlConnection(Connection);
            ketNoi();
            DataTable dt = new DataTable();
            try
            {
                SqlDataAdapter sqlda = new SqlDataAdapter(sql, sqlcon);
                sqlda.Fill(dt);
                dongKetNoi();
            }
            catch
            {
                MessageBox.Show("Lỗi truy vấn dữ liệu ", "Cảnh báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
            }

            return dt;
            
        }
        
        public int Update(string sql)
        {
           try
            {
                //string Connection = GetConnectionString();
                // sqlcon = new SqlConnection(Connection);
                ketNoi();
                SqlCommand sqlcom = new SqlCommand(sql,sqlcon);               
                return sqlcom.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Lỗi truy cập không trả về được dữ liệu", "Cảnh báo");
                return -1;
            }

        }

        public SqlCommand SqlCommand(string sql)
        {
            try
            {
                //string Connection = GetConnectionString();
                //sqlcon = new SqlConnection(Connection);
                ketNoi();
                SqlCommand command = new SqlCommand(sql, sqlcon);
                return command;
            }
            catch
            {
                MessageBox.Show("Lỗi truy cập không trả về được dữ liệu", "Cảnh báo");
                return null;
            }
            
        }

    }
}
