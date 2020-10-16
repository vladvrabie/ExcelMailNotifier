using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace MailingWindowsService
{
    public partial class MailService : ServiceBase
    {
        public MailService()
        {
            InitializeComponent();

            //ServiceName = "...";
            //eventLogger.Source = "...";
            //eventLogger.Log = "...";

            if (!EventLog.SourceExists(eventLogger.Source))
            {
                EventLog.CreateEventSource(eventLogger.Source, eventLogger.Log);
            }
        }

        protected override void OnStart(string[] args)
        {
            eventLogger.WriteEntry("In OnStart");
        }

        protected override void OnStop()
        {
            eventLogger.WriteEntry("In OnStop");
        }
    }
}
