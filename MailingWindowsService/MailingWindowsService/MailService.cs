using System;
using System.Diagnostics;
using System.ServiceProcess;
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
            eventLogger.LogI("In SetupTimerForExcelRead.");
        }

        private void SetupTimerForSendEmail()
        {
            timer.Stop();
            timer.Interval = 60000; // 1 minute
            timer.AutoReset = true;
            timer.Elapsed += SendEmail;
            timer.Start();
            eventLogger.LogI("In SetupTimerForSendEmail.");
        }

        private void SetupTimerToWait24H()
        {
            timer.Stop();
            timer.Interval = 24.0/*h/day*/ * 60/*min/h*/ * 60/*s/min*/ * 1000/*ms/s*/;  // 1 day in ms
            //timer.Interval = 1.0/*min*/ * 60/*s/min*/ * 1000/*ms/s*/;  // 1 min in ms
            timer.AutoReset = false;
            timer.Elapsed += After24HElapsed;
            timer.Start();
            eventLogger.LogI("In SetupTimerToWait24H.");
        }

        private void After24HElapsed(object sender, ElapsedEventArgs args)
        {
            timer.Elapsed -= After24HElapsed;
            currentDate = DateTime.Today;
            SetupTimerForExcelRead();
        }

        private void GetExcelData(object sender, ElapsedEventArgs args)
        {
            if (TryReadExcel())
            {
                timer.Elapsed -= GetExcelData;
                if (ShouldSendEmail())
                {
                    //eventLogger.LogI("Should send email true");
                    SetupTimerForSendEmail();
                }
                else
                {
                    //eventLogger.LogI("Should send email false");
                    SetupTimerToWait24H();
                }
            }
            else
            {
                //eventLogger.LogI("Try read excel false");
                // SetupTimerForExcelRead();
            }
        }

        private bool TryReadExcel()
        {
            var parameters = new ExcelReader.AppConfigReader()
            {
                logger = eventLogger
            }.GetExcelReaderParameters();

            //var nill = "null";
            //eventLogger.LogI($"path: {parameters.path ?? nill}");
            //eventLogger.LogI($"sheetsNames: {parameters.sheetsNames?.ToString() ?? nill}");
            //eventLogger.LogI($"sheetsIndexes: {parameters.sheetsIndexes?.ToString() ?? nill}");
            //eventLogger.LogI($"headerRow: {parameters.headerRow}");
            //eventLogger.LogI($"columnsToCheckDate: {parameters.columnsToCheckDate?.ToString() ?? nill}");
            //eventLogger.LogI($"columnsIndexesToCheckDate: {parameters.columnsIndexesToCheckDate?.ToString() ?? nill}");
            //eventLogger.LogI($"dateFormats: {parameters.dateFormats?.ToString() ?? nill}");
            //eventLogger.LogI($"daysUntilExpirationCheck: {parameters.daysUntilExpirationCheck?.ToString() ?? nill}");
            //eventLogger.LogI($"columnsToEmail: {parameters.columnsToEmail?.ToString() ?? nill}");
            //eventLogger.LogI($"columnsIndexesToEmail: {parameters.columnsIndexesToEmail?.ToString() ?? nill}");

            var excelReader = new ExcelReader.NPOIExcelReader(parameters)
            {
                logger = eventLogger
            };

            try
            {
                excelData = excelReader.Get();
                return excelData != null;
            }
            catch (Exception ex)
            {
                eventLogger.LogE($"Exception in ExcelReader.Get\nMessage: {ex.Message}\nSource: {ex.Source}\nStack trace: {ex.StackTrace}");
                return false;
            }
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
                currentDate = DateTime.Today;
                SetupTimerForExcelRead();
            }
        }

        private bool TrySendEmail()
        {
            var parameters = new EmailSender.AppConfigReader() 
            {
                logger = eventLogger
            }.GetEmailSenderParameters();
            var emailSender = new EmailSender.MailKitEmailSender(parameters)
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
