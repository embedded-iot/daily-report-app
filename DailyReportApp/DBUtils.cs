using System;
using System.Collections.Generic;
using System.Text;
using Oracle.DataAccess.Client;

namespace DailyReportApp
{
    class DBUtils
    {
        public static OracleConnection GetDBConnection()
        {
            string host = "127.0.0.1";
            int port = 1521;
            string sid = "xe";
            string user = "Project1";
            string password = "Project1";
            return DBOracleUtils.GetDBConnection(host, port, sid, user, password);
        }
    }
}
