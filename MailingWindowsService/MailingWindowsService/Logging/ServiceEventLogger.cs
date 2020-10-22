using System.Diagnostics;

namespace MailingWindowsService.Logging
{
    class ServiceEventLogger : EventLog, ILogger
    {
        private int messageId = 0;

        public void LogE(string message) => WriteEntry(message, EventLogEntryType.Error, ++messageId);

        public void LogI(string message) => WriteEntry(message, EventLogEntryType.Information, ++messageId);

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
