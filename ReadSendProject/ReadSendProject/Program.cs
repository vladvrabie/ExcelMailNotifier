using ReadSendProject.ExcelReader;
using ReadSendProject.Logging;
using System;
using System.Collections.Generic;

namespace ReadSendProject
{
    class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("running :)");
            var readerParams = new ExcelReaderParameters()
            {
                path = @"...",
                sheetsNames = new List<string> { "..." },
                headerRow = 2, 
                daysUntilExpirationCheck = new List<int> { 1, 2 },
                columnsIndexesToCheckDate = new List<int> { 5, 6, 7 },
                dateFormats = new List<string> { "dd/MM/yyyy", "dd.MM.yyyy" },
                columnsIndexesToEmail = new List<int> { 2, 3, 5, 6, 7 },
            };

            var reader = new NPOIExcelReader(readerParams)
            {
                logger = new ConsoleLogger()
            };

            var r = reader.Get();
            foreach (var row in r)
            {
                foreach (var cell in row)
                {
                    Console.Write($"{cell}\t");
                }
                Console.WriteLine();
            }


            //ExcelReaderParameters readerParameters = new ExcelReaderParameters()
            //{
            //    path = @"...",
            //    sheetsNames = new List<string> { "..." },
            //    //sheetsIndexes = new List<int> { 2, 1 },
            //    headerRow = 3,
            //    columnsToCheckDate = new List<string> { "E", "F", "G" },
            //    dateFormats = new List<string> { "dd/MM/yyyy", "dd.MM.yyyy" },
            //    daysUntilExpirationCheck = new List<int> { 6, 9 },
            //    columnsToEmail = new List<string> { "A", "B", "E", "F", "G" },
            //};

            //MSInteropExcelReader reader = new MSInteropExcelReader(readerParameters)
            //{
            //    logger = new Logging.ConsoleLogger()
            //};
            //var results = reader.Get();


            //EmailSenderParameters emailParameters = new EmailSenderParameters()
            //{
            //    senderEmail = "...",
            //    senderEmailPassword = "...",
            //    receiverEmails = new List<string> { "..." }
            //};

            //MailKitEmailSender sender = new MailKitEmailSender(emailParameters)
            //{
            //    logger = new Logging.ConsoleLogger()
            //};
            //sender.TrySendEmail(results);


            //foreach (var row in results)
            //{
            //    foreach (var cell in row)
            //    {
            //        Console.Write($"{cell}\t");
            //    }
            //    Console.WriteLine();
            //}

            //Console.WriteLine();
            //Console.WriteLine();
            //Console.WriteLine(StringMatrixConverter.ToHtmlTable(results));

            //Console.WriteLine();
            //Console.WriteLine();
            //Console.WriteLine(StringMatrixConverter.ToPlainTextTable(results));
        }
    }
}
