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
            string[,] data = { };
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
                    int countColumn = columnNames.Length > 0 ? columnNames.Length : reader.FieldCount;
                    data = new string[countRows, countColumn];
                    while (reader.Read())
                    {
                        for (int index = 0; index < countColumn; index++)
                        {
                            int empNameIndex = columnNames.Length > 0 ? reader.GetOrdinal(columnNames[index]) : index;
                            string empName = reader.GetValue(empNameIndex).ToString();
                            data[rowIndex, index] = empName;

                        }
                        rowIndex++;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("QueryAllData Error: " + ex);
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
            Console.WriteLine("Write Data In ==>" + " " + excelPath);
            OracleConnection conn = DBOracleUtils.GetDBConnection(host, port, sid, user, password);
            //Console.WriteLine("Connecting" + DateTime.Now);
            try
            {
                conn.Open();
                // Console.WriteLine("Read columnNames" + DateTime.Now);
                //string[] columnNames = ExcelUtils.ReadRow(excelPath, sheetIndex, 1);
                string[] columnNames = {}; // Nếu column name = {}, thì sẽ ghi toàn bộ giữ liệu từ DB tới excel
                //Console.WriteLine("QueryAllData" + DateTime.Now);
                string[,] data = QueryAllData(conn, columnNames, tableName);
                //Console.WriteLine("Rows: " + data.GetLength(0));
                //Console.WriteLine("Columns: " + data.GetLength(1));
                //Console.WriteLine("AppendRows" + DateTime.Now);
                //ExcelUtils.AppendRowsNew(excelPath, sheetIndex, data, true);
                ExcelUtils.AppendRowsNew(excelPath, sheetIndex, data, true);
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
            Console.WriteLine("Successful!");
        }
        static void Main(string[] args)
        {   // chạy theo schedule 
            string svn = "D:\\Code\\SVN LOCAL\\test.xlsx";
            bool isUpdated = SharpSVNAgent.SvnCommit(svn, "Update Today");
            //if (!isUpdated)
            //{
            //    Console.WriteLine("Update failure");
            //    return;
            //}
            Console.WriteLine("DAILY REPORT TRANFERSJOB");
            
            CronTabHandler("127.0.0.1", 1521, "xe", "Project2", "Project2", "na_r_acs_tranfertimejob_vw", svn, 1);
            //CronTabHandler("127.0.0.1", 1521, "xe", "Project2", "Project2", "na_r_acs_tranfertimejob_vw", svn, 1);
            //CronTabHandler("127.0.0.1", 1521, "xe", "Project1", "Project1", "hssv", "D:\\Code\\Svn\\test.xlsx", 2);
            //CronTabHandler("127.0.0.1", 1521, "xe", "Project2", "Project2", "hssv", "D:\\Code\\Svn\\test.xlsx", 2);
            //Console.WriteLine("Commiting...");
            //return;
            bool isCommited = SharpSVNAgent.SvnCommit(svn, "Commit Today");
            if (isCommited)
            {
                Console.WriteLine("Commited sucessful");
            }
            else
            {
                Console.WriteLine("Commited failure");
            }
            Console.ReadKey();
        }
    }
}