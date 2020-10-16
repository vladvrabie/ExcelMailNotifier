using System;
using System.Globalization;

namespace ReadSendProject.Logging
{
    class ConsoleLogger : AbstractLogger
    {
        public override void Log(string message, MessageType type)
        {
            var datetime = DateTime.Now.ToString(CultureInfo.CreateSpecificCulture("ro-RO"));
            Console.WriteLine($"[{datetime}] {type}: {message}");
        }
    }
}
