namespace ReadSendProject.Logging
{
    interface ILogger
    {
        void Log(string message, MessageType type);

        void LogI(string message);

        void LogE(string message);
    }
}
