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
        
        private static string[,] QueryAllData(OracleConnection conn, string[] columnNames, string tableName)
        {
            string[,] data = {};
            int rowIndex = 0;
            DbDataReader reader = DBOracleUtils.ExecuseQuery(conn, "SELECT * from " + tableName);
            DbDataReader readerCount = DBOracleUtils.ExecuseQuery(conn, "SELECT COUNT(*) from " + tableName);
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

        private static void CronTabHandler(string host, int port, string sid, string user, string password, string tableName, string excelPath, int sheetIndex)
        {
            Console.Write("Generate excel file ...", excelPath);
            OracleConnection conn = DBOracleUtils.GetDBConnection(host, port, sid, user, password);
            try
            {
                conn.Open();
                string[] columnNames = ExcelUtils.ReadRow(excelPath, sheetIndex, 1);
                string[,] data = QueryAllData(conn, columnNames, tableName);
                ExcelUtils.AppendRows(excelPath, sheetIndex, data, true);
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

        private static void Stick()
        {
            Console.WriteLine("Stick");
        }


        static void Main(string[] args)
        {
            //for (int connectIndex = 0; connectIndex < connectList.Length; connectIndex++)
            //{
               
            //    CronTabHandler("127.0.0.1", 1521, "xe", "Project1", "Project1",  "HSSV", "D:\\Code\\test.xlsx");  // chạy trực tiếp app dùng schedule bên ngoàii
            //}

            CronTabHandler("127.0.0.1", 1521, "xe", "Project1", "Project1", "HSSV", "D:\\Code\\test.xlsx", 2);  // chạy trực tiếp app dùng schedule bên ngoàii

            //CronTabHandler();
            //CronTabHandler1();

            return;
            // chạy schedule trong service, check hàm CronTabHandler
            string cronTabExpression = "* * * * *";  // Refer the doctument https://crontab.guru/
            CronTab cronTab = new CronTab(cronTabExpression, Stick);
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