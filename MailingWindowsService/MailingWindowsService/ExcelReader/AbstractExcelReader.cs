using MailingWindowsService.Logging;
using StringMatrix = System.Collections.Generic.List<System.Collections.Generic.List<string>>;

namespace MailingWindowsService.ExcelReader
{
    abstract class AbstractExcelReader
    {
        public ILogger logger;

        protected readonly ExcelReaderParameters excelParameters;

        public AbstractExcelReader(ExcelReaderParameters parameters)
        {
            excelParameters = parameters;
        }

        public abstract StringMatrix Get();
    }
}
