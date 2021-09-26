using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Common;
using Oracle.DataAccess.Client;


namespace DailyReportApp
{
    static class Program
    {
        static void Main(string[] args)
        {
            OracleConnection conn = DBUtils.GetDBConnection();
            try
            {
                conn.Open();
                Console.WriteLine(conn.ConnectionString + "Successful!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                conn.Close();
                conn.Dispose();
            }

            Console.Read();
        }

        private static void QueryEmployee(OracleConnection conn)
        {
            string sql = "SELECT * FROM na_r_acs_tranfertimejob_vw";
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.CommandText = sql;
            using (DbDataReader reader = cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        int a = 0;
                    }
                }
            }

        }
    }
}