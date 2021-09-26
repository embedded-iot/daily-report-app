using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Topshelf;

namespace DailyReportApp
{
    class WindowsService
    {
        private CronTab _cronTabInstance;
        private string _serviceName;
        private string _displayName;
        private string _description;


        public WindowsService(CronTab cronTabInstance, string serviceName, string displayName, string description)
        {
            _cronTabInstance = cronTabInstance;
            _serviceName = serviceName;
            _displayName = displayName;
            _description = description;
        }

        public void Start()
        {
            var exitCode = HostFactory.Run(x =>
            {
                x.Service<CronTab>(s =>
                {
                    s.ConstructUsing(() => _cronTabInstance);
                    s.WhenStarted(_cronTabInstance => _cronTabInstance.Start());
                    s.WhenStopped(_cronTabInstance => _cronTabInstance.Stop());
                });

                x.RunAsLocalSystem();

                x.SetServiceName(_serviceName);
                x.SetDisplayName(_displayName);
                x.SetDescription(_description);
            });

            int exitCodeValue = (int)Convert.ChangeType(exitCode, exitCode.GetTypeCode());
            Environment.ExitCode = exitCodeValue;
        }
    }
}
