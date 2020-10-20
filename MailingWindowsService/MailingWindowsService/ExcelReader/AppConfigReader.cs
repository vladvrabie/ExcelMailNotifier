﻿using MailingWindowsService.Logging;
using System;
using System.Configuration;
using System.Linq;

namespace MailingWindowsService.ExcelReader
{
    class AppConfigReader
    {
        public ILogger logger;

        private static readonly string PATH = "excelPath";
        private static readonly string SHEETS_NAMES = "sheetsNames";
        private static readonly string SHEETS_INDEXES = "sheetsIndexes";
        private static readonly string HEADER_ROW = "headerRow";
        private static readonly string COLUMNS_TO_CHECK = "columnsToCheckDate";
        private static readonly string DATE_FORMATS = "dateFormats";
        private static readonly string DAYS_UNTIL_EXPIRATION = "daysUntilExpirationCheck";
        private static readonly string COLUMNS_TO_EMAIL = "columnsToEmail";
        
        public ExcelReaderParameters GetExcelReaderParameters()
        {
            try
            {
                var appSettings = ConfigurationManager.AppSettings;
                if (appSettings.Count == 0)
                {
                    throw new ConfigurationErrorsException();
                }

                return new ExcelReaderParameters()
                {
                    path = appSettings[PATH],
                    sheetsNames = appSettings[SHEETS_NAMES]?.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList(),
                    sheetsIndexes = appSettings[SHEETS_INDEXES]?.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList().ConvertAll(int.Parse),
                    headerRow = int.Parse(appSettings[HEADER_ROW]),
                    columnsToCheckDate = appSettings[COLUMNS_TO_CHECK]?.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList(),
                    dateFormats = appSettings[DATE_FORMATS]?.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList(),
                    daysUntilExpirationCheck = appSettings[DAYS_UNTIL_EXPIRATION]?.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList().ConvertAll(int.Parse),
                    columnsToEmail = appSettings[COLUMNS_TO_EMAIL]?.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList(),
                };
            }
            catch (Exception ex)
            {
                logger.LogE($"Error in retrieving excel settings for the app; Message: {ex.Message}");
                return new ExcelReaderParameters();
            }
        }
    }
}
