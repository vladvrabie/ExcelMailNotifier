using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadSendProject
{
    class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("running :)");

            ExcelReaderParameters readerParameters = new ExcelReaderParameters()
            {
                path = @"...",
                sheetsNames = new List<string> { "..." },
                //sheetsIndexes = new List<int> { 2, 1 },
                headerRow = 3,
                columnsToCheckDate = new List<string> { "E", "F", "G" },
                dateFormats = new List<string> { "dd/MM/yyyy", "dd.MM.yyyy" },
                daysUntilExpirationCheck = new List<int> { 6, 9 },
                columnsToEmail = new List<string> { "A", "B", "E", "F", "G" },
            };

            ExcelReader reader = new ExcelReader(readerParameters)
            {
                logger = new Logging.ConsoleLogger()
            };
            var results = reader.Get();


            EmailSenderParameters emailParameters = new EmailSenderParameters()
            {
                senderEmail = "...",
                senderEmailPassword = "...",
                receiverEmails = new List<string> { "..." }
            };

            EmailSender sender = new EmailSender(emailParameters)
            {
                logger = new Logging.ConsoleLogger()
            };
            sender.TrySendEmail(results);


            foreach (var row in results)
            {
                foreach (var cell in row)
                {
                    Console.Write($"{cell}\t");
                }
                Console.WriteLine();
            }

            //Console.WriteLine();
            //Console.WriteLine();
            //Console.WriteLine(StringMatrixConverter.ToHtmlTable(results));

            //Console.WriteLine();
            //Console.WriteLine();
            //Console.WriteLine(StringMatrixConverter.ToPlainTextTable(results));
        }
    }
}
