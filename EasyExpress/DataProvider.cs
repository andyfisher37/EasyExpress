/// Class Provider for working with DATABASES

using System;
using System.Data;
using System.Data.SqlClient;

namespace EasyExpress
{
    class DataProvider
    {
        // Request SQL data and return DataTable result
        public static DataTable _getDataSQL(string constr, string cmd)
        {
            try
            {
                SqlConnection Conn = new SqlConnection(constr);
                Conn.Open();
                SqlTransaction tr = Conn.BeginTransaction();
                DataTable dt = new DataTable();
                SqlCommand cm = new SqlCommand(cmd, Conn) { Transaction = tr };
                cm.CommandType = CommandType.Text;
                SqlDataReader dr = cm.ExecuteReader();
                dt.Load(dr);
                tr.Commit();
                Conn.Close();
                dr.Dispose();
                cm.Dispose();
                tr.Dispose();
                return dt;
            } catch (Exception ex)
            { 
                return null;
            }

        }

		// Request SQL data and return one integer result (simple)
		public static int _getDataSQLs(string constr, string cmd)
        {
            try
            {
                SqlConnection Conn = new SqlConnection(constr);
                Conn.Open();
                SqlCommand cm = new SqlCommand(cmd, Conn);
                cm.CommandType = CommandType.Text;
                int res = Convert.ToInt16(cm.ExecuteScalar());
                Conn.Close();
                cm.Dispose();
                return res;
            }
            catch(Exception ex) { return 0; }
        }

        // Insert into Table SQL data
        public static int _insDataSQL(string constr, string cmd)
        {
            try
            {
                int res = 0;
                SqlConnection Conn = new SqlConnection(constr);
                Conn.Open();
                SqlTransaction tr = Conn.BeginTransaction();
                SqlCommand cm = new SqlCommand(cmd, Conn) { Transaction = tr };
                cm.CommandType = CommandType.Text;
                res = cm.ExecuteNonQuery();
                tr.Commit();
                Conn.Close();
                cm.Dispose();
                tr.Dispose();
                return res;
            }
            catch(Exception ex) { return 0; }
        }



    }
}
