using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using NCrontab;

namespace DailyReportApp
{
    class CronTab
    {
        private string _cronTabExpression = "* * * * *";
        private Action _cronTabCallback;
        private Timer _timer;
        public CronTab(string cronTabExpression, Action cronTabCallback)
        {
            _cronTabExpression = cronTabExpression ;
            _cronTabCallback = cronTabCallback;
            _timer = new Timer(1000);
            _timer.Elapsed += _ElapsedEventHandler;
        }

        private void _ElapsedEventHandler(object sender, ElapsedEventArgs e)
        {
            // Handler function will be triggerred every seconds
            CrontabSchedule cronTabChedule = CrontabSchedule.Parse(_cronTabExpression);
            DateTime _now = DateTime.Now;
            DateTime _next = cronTabChedule.GetNextOccurrence(_now.AddSeconds(-1));
            if (_now.ToString() == _next.ToString())
            {
                _cronTabCallback();
            }
        }

        public void Start()
        {
            _timer.Start();
        }

        public void Stop()
        {
            _timer.Stop();
        }
    }
}
