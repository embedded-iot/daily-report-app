using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SharpSvn;

namespace DailyReportApp
{
    public static class SharpSVNAgent
    {
        public static bool CheckIsSvnDir(string path)
        {
            SvnClient client = new SvnClient();
            return false;
        }

        public static bool SvnUpdate(string path)
        {
            SvnClient client = new SvnClient();
            return client.Update(path);
        }

        public static bool SvnCommit(string path, string mess)
        {
            SvnCommitArgs args = new SvnCommitArgs();
            args.LogMessage = mess;
            SvnClient client = new SvnClient();
            return client.Commit(path, args);
        }
    }
}