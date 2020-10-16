using System.Diagnostics;

namespace MailingWindowsService.Logging
{
    class ServiceEventLogger : EventLog, ILogger
    {
        public void LogE(string message) => WriteEntry(message, EventLogEntryType.Error);

        public void LogI(string message) => WriteEntry(message, EventLogEntryType.Information);

        void ILogger.Log(string message, MessageType type)
        {
            switch (type)
            {
                case MessageType.INFO:
                    LogI(message);
                    break;
                case MessageType.ERROR:
                    LogE(message);
                    break;
                default:
                    break;
            }
        }
    }
}
