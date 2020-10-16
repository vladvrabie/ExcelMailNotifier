using System.Collections.Generic;

namespace MailingWindowsService.ExcelReader
{
    struct ExcelReaderParameters
    {
        /// <summary>
        /// Path to the excel to open; better be full path.
        /// </summary>
        public string path;

        /// <summary>
        /// Work sheets to be processed, found by name.
        /// If indexes are also used, duplicates are not eliminated.
        /// </summary>
        public List<string> sheetsNames;

        /// <summary>
        /// Work sheets to be processed, found by index (starting at 1).
        /// If names are also used, duplicates are not eliminated.
        /// </summary>
        public List<int> sheetsIndexes;

        /// <summary>
        /// Index (from 1) of Excel table header to process, assuming all
        /// given sheets have same header row at the same row index.
        /// It will be used as header row in the email.
        /// Date checks start below this row.
        /// </summary>
        public int headerRow;

        /// <summary>
        /// List of column addresses that will be checked for expiration date.
        /// </summary>
        public List<string> columnsToCheckDate;

        /// <summary>
        /// List of date formats to use if the dates in the sheets are text.
        /// </summary>
        public List<string> dateFormats;

        /// <summary>
        /// List with number of future days to check until expiration.
        /// </summary>
        public List<int> daysUntilExpirationCheck;

        /// <summary>
        /// List of column addresses that will be included in the email.
        /// The columns checked for date are NOT included.
        /// First column in email will contain the sheet and row index of expired row.
        /// </summary>
        public List<string> columnsToEmail;
    }
}
