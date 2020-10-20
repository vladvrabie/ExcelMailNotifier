using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using StringMatrix = System.Collections.Generic.List<System.Collections.Generic.List<string>>;


namespace MailingWindowsService
{
    public partial class MailService : ServiceBase
    {
        private readonly Timer timer = new Timer();
        private StringMatrix excelData = null;
        private DateTime currentDate = DateTime.Today;

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
            eventLogger.LogI("In OnStart.");
            SetupTimerForExcelRead();
        }

        protected override void OnStop()
        {
            eventLogger.LogI("In OnStop.");
            timer.Stop();
        }

        private void SetupTimerForExcelRead()
        {
            timer.Stop();
            timer.Interval = 60000; // 1 minute
            timer.AutoReset = true;
            timer.Elapsed += GetExcelData;
            timer.Start();
        }

        private void SetupTimerForSendEmail()
        {
            timer.Stop();
            timer.Interval = 60000; // 1 minute
            timer.AutoReset = true;
            timer.Elapsed += SendEmail;
            timer.Start();
        }

        private void SetupTimerToWait24H()
        {
            timer.Stop();
            timer.Interval = 24/*h/day*/ * 60/*min/h*/ * 60/*s/min*/ * 1000/*ms/s*/;  // 1 day in ms
            timer.AutoReset = false;
            timer.Elapsed += After24HElapsed;
            timer.Start();
            eventLogger.LogI("No email today.");
        }

        private void After24HElapsed(object sender, ElapsedEventArgs args)
        {
            timer.Elapsed -= After24HElapsed;
            SetupTimerForExcelRead();
        }

        private void GetExcelData(object sender, ElapsedEventArgs args)
        {
            if (TryReadExcel())
            {
                timer.Elapsed -= GetExcelData;
                if (ShouldSendEmail())
                {
                    SetupTimerForSendEmail();
                }
                else
                {
                    SetupTimerToWait24H();
                }
            }
            else
            {
                // SetupTimerForExcelRead();
            }
        }

        private bool TryReadExcel()
        {
            var parameters = new ExcelReader.AppConfigReader() 
            { 
                logger = eventLogger 
            }.GetExcelReaderParameters();
            var excelReader = new ExcelReader.ExcelReader(parameters) 
            { 
                logger = eventLogger 
            };
            excelData = excelReader.Get();
            return excelData != null;
        }

        bool ShouldSendEmail()
        {
            // First row is the header.
            // If we have 2 rows, we have at least 1 row with data to send.
            return excelData.Count >= 2;
        }

        private void SendEmail(object sender, ElapsedEventArgs args)
        {
            if (ADayPassed() == false)
            {
                if (TrySendEmail())
                {
                    timer.Elapsed -= SendEmail;
                    SetupTimerToWait24H();
                }
                else
                {
                    // SetupTimerForSendEmail();
                }
            }
            else
            {
                timer.Elapsed -= SendEmail;
                SetupTimerForExcelRead();
            }
        }

        private bool TrySendEmail()
        {
            var parameters = new EmailSender.AppConfigReader() 
            {
                logger = eventLogger
            }.GetEmailSenderParameters();
            var emailSender = new EmailSender.EmailSender(parameters)
            {
                logger = eventLogger
            };
            return emailSender.TrySendEmail(excelData);
        }

        private bool ADayPassed()
        {
            var timeDifference = DateTime.Today - currentDate;
            return timeDifference.Days != 0;
        }
    }
}
