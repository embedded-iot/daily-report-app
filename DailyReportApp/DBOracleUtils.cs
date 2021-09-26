using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.DataAccess.Client;

namespace DailyReportApp
{
    class DBOracleUtils
    {
        public static OracleConnection GetDBConnection(string host, int port, string sid, string user, string password)
        {
            Console.WriteLine("Getting Connection ...");
            // Connection String để kết nối trực tiếp tới Oracle.
            string connString = "Data Source=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = "
            + host + ")(PORT = " + port + "))(CONNECT_DATA = (SERVER = DEDICATED)(SERVICE_NAME = "
            + sid + ")));Password=" + password + ";User ID=" + user;

            OracleConnection conn = new OracleConnection
            {
                ConnectionString = connString
            };
            return conn;
        }

    }
}