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
        private static string[] columnNames = { "FLOOR", "LONGDATE", "SHORTDATE" };
        private static OracleConnection GetDBConnection()
        {
            string host = "127.0.0.1";
            int port = 1521;
            string sid = "xe";
            string user = "Project1";
            string password = "Project1";
            return DBOracleUtils.GetDBConnection(host, port, sid, user, password);
        }

        private static string[,] QueryAllData(OracleConnection conn)
        {
            string[,] data;
            int rowIndex = 0;
            DbDataReader reader = DBOracleUtils.ExecuseQuery(conn, "SELECT * from na_r_acs_tranfertimejob_vw");
            using (reader)
            {
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        for (int index = 0; index < columnNames.Length; index++)
                        {
                            int empNameIndex = reader.GetOrdinal(columnNames[index]);
                            string empName = reader.GetString(empNameIndex);
                            data[rowIndex, index] = empName;
                            
                        }

                    }
                }
            }
            return data;
        }

        private static void CronTabHandler()
        {
            OracleConnection conn = GetDBConnection();
            try
            {
                conn.Open();
                string[,] data = QueryAllData(conn);
                string[,] data1 = {
                                 {"1", "2"},
                                 {"3", "4"},
                                 {"5", "6"}
                             };
                ExcelUtils.AppendRows("F:\\test.xlsx", data);
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
        }


        static void Main(string[] args)
        {
            string cronTabExpression = "* * * * *";  // Refer the doctument https://crontab.guru/
            CronTab cronTab = new CronTab(cronTabExpression, CronTabHandler);
            WindowsService service = new WindowsService(
                cronTab,
                "DailyReportService",
                "Daily Report Service",
                "This service will generate the report excel file daily."
                );
            service.Start();
        }
    }
}