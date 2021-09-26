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
        
        private static OracleConnection GetDBConnection()
        {
            string host = "127.0.0.1";
            int port = 1521;
            string sid = "xe";
            string user = "Project1";
            string password = "Project1";
            return DBOracleUtils.GetDBConnection(host, port, sid, user, password);
        }

        private static string[,] QueryAllData(OracleConnection conn, string[] columnNames)
        {
            string[,] data = {};
            int rowIndex = 0;
            DbDataReader reader = DBOracleUtils.ExecuseQuery(conn, "SELECT * from na_r_acs_tranfertimejob_vw");
            DbDataReader readerCount = DBOracleUtils.ExecuseQuery(conn, "SELECT COUNT(*) from na_r_acs_tranfertimejob_vw");
            int countRows = 0;
            try
            {
                if (readerCount.HasRows)
                {
                    readerCount.Read();
                    countRows = readerCount.GetInt32(0);
                }
                if (reader.HasRows)
                {
                    data = new string[countRows, columnNames.Length];
                    while (reader.Read())
                    {
                        for (int index = 0; index < columnNames.Length; index++)
                        {
                            int empNameIndex = reader.GetOrdinal(columnNames[index]);
                            string empName = reader.GetValue(empNameIndex).ToString();
                            data[rowIndex, index] = empName;

                        }
                        rowIndex++;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex);
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                reader.Dispose();
            }
            return data;
        }

        private static void CronTabHandler()
        {
            Console.WriteLine("Generate excel file...");
            OracleConnection conn = GetDBConnection();
            string path = "F:\\test.xlsx";
            try
            {
                conn.Open();
                string[] columnNames = ExcelUtils.ReadRow(path, 1);
                string[,] data = QueryAllData(conn, columnNames);
                ExcelUtils.AppendRows(path, data);
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
            Console.WriteLine("Done!");
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