namespace MailingWindowsService.Logging
{
    abstract class AbstractLogger : ILogger
    {
        public abstract void Log(string message, MessageType type);

        public void LogE(string message) => Log(message, MessageType.ERROR);

        public void LogI(string message) => Log(message, MessageType.INFO);
    }
}
