using MailingWindowsService.Logging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using StringMatrix = System.Collections.Generic.List<System.Collections.Generic.List<string>>;

namespace MailingWindowsService.ExcelReader
{
    class ExcelReader
    {
        public ILogger logger;

        private readonly ExcelReaderParameters excelParams;

        private Excel.Application application;
        private Excel.Workbook workbook;
        private readonly List<Excel._Worksheet> sheets;
        private readonly List<Excel.Range> ranges;

        public ExcelReader(ExcelReaderParameters parameters)
        {
            excelParams = parameters;

            int sheetsNamesCount = excelParams.sheetsNames?.Count ?? 0;
            int sheetsIndexesCount = excelParams.sheetsIndexes?.Count ?? 0;
            sheets = new List<Excel._Worksheet>(sheetsNamesCount + sheetsIndexesCount);
            ranges = new List<Excel.Range>(sheets.Capacity);
        }

        private bool TryOpenExcel()
        {
            try
            {
                application = new Excel.Application();
                workbook = application.Workbooks.Open(excelParams.path, ReadOnly: true);
                return true;
            }
            catch (COMException ex)
            {
                logger?.LogE($"Error while trying to open excel; Error code: {ex.ErrorCode} ; Message: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                logger?.LogE($"Exception in TryOpenExcel: {ex.Message}");
                return false;
            }
        }

        private void CloseExcel()
        {
            // cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Close Excel workbook
            if (workbook != null)
            {
                workbook.Close();
                Marshal.ReleaseComObject(workbook);
            }

            // Close Excel application
            if (application != null)
            {
                application.Quit();
                Marshal.ReleaseComObject(application);
            }
        }

        private bool TryGetSheetsAndRanges()
        {
            if (excelParams.sheetsNames.IsNullOrEmpty() == false)
            {
                foreach (string sheetName in excelParams.sheetsNames)
                {
                    try
                    {
                        Excel._Worksheet sheet = workbook.Sheets[sheetName];
                        sheets.Add(sheet);
                        ranges.Add(sheet.UsedRange);
                    }
                    catch (COMException)
                    {
                        logger?.LogE($"Incorrect name given for a worksheet: {sheetName}");
                        continue;
                    }
                }
            }

            if (excelParams.sheetsIndexes.IsNullOrEmpty() == false)
            {
                foreach (int index in excelParams.sheetsIndexes)
                {
                    try
                    {
                        Excel._Worksheet sheet = workbook.Sheets[index];
                        sheets.Add(sheet);
                        ranges.Add(sheet.UsedRange);
                    }
                    catch (COMException)
                    {
                        logger?.LogE($"Incorrect index given for a worksheet: {index}");
                        continue;
                    }
                }
            }

            if (sheets.Count == 0)
            {
                logger?.LogI("No sheets to process");
                return false;
            }

            return true;
        }

        private void ReleaseRangesAndSheets()
        {

            for (int i = 0; i < ranges.Count; i++)
            {
                Marshal.ReleaseComObject(ranges[i]);
                Marshal.ReleaseComObject(sheets[i]);
            }

            ranges.Clear();
            sheets.Clear();
        }

        public StringMatrix Get()
        {
            if (TryOpenExcel())
            {
                StringMatrix result = null;

                if (TryGetSheetsAndRanges())
                {
                    int estimatedCapacityUpperLimit = 10;
                    result = new StringMatrix(estimatedCapacityUpperLimit)
                    {
                        GetHeaderColumnsToEmail()  // always returns a list with at least 1 element
                    };
                    result.AddRange(ProcessCells()); // can return empty list

                    ReleaseRangesAndSheets();
                }

                CloseExcel();

                return result;
            }

            return null;
        }

        private StringMatrix ProcessCells()
        {
            StringMatrix result = new StringMatrix();

            for (int sheetIndex = 0; sheetIndex < sheets.Count; sheetIndex++)
            {
                var sheet = sheets[sheetIndex];
                int startRowIndex = excelParams.headerRow + 1;
                int endRowIndex;
                {
                    var range = ranges[sheetIndex];
                    int nbRows = range.Rows.Count;
                    int nbCols = range.Columns.Count;
                    Excel.Range lastUsedCell = range[nbRows, nbCols];
                    endRowIndex = lastUsedCell.Row;
                }

                for (int rowIndex = startRowIndex; rowIndex <= endRowIndex; rowIndex++)
                {
                    if (IsRowExpired(sheet, rowIndex))
                    {
                        var cells = new List<string>((excelParams.columnsToEmail?.Count ?? 0) + 1)
                        {
                            $"{sheet.Name} {rowIndex}"
                        };
                        cells.AddRange(GetColumnsToEmailAtRow(sheet, rowIndex));

                        result.Add(cells);
                    }
                }
            }

            return result;
        }

        private bool IsRowExpired(Excel._Worksheet sheet, int rowIndex)
        {
            if (excelParams.columnsToCheckDate != null)
            {
                foreach (string columnIndex in excelParams.columnsToCheckDate)
                {
                    Excel.Range cell = sheet.Cells[rowIndex, columnIndex];

                    if (cell.Value == null)
                    {
                        continue;
                    }

                    if (TryGetDate(cell.Value, out DateTime cellDate))
                    {
                        try
                        {
                            TimeSpan timeDifference = cellDate - DateTime.Today;
                            if (excelParams.daysUntilExpirationCheck.Contains(timeDifference.Days))
                            {
                                return true;
                            }
                        }
                        catch (ArgumentOutOfRangeException)
                        {
                            continue;
                        }
                    }
                }
            }
            return false;
        }

        private bool TryGetDate(dynamic cellValue, out DateTime date)
        {
            if (cellValue is DateTime)
            {
                date = cellValue;
                return true;
            }
            else if (cellValue is string)
            {
                return DateTime.TryParseExact(
                    cellValue,
                    excelParams.dateFormats.ToArray(),
                    null,
                    DateTimeStyles.None,
                    out date
                );
            }
            else if (cellValue is double)
            {
                try
                {
                    date = DateTime.FromOADate(cellValue); // maybe??
                    return true;
                }
                catch (ArgumentException)
                {
                    // Nothing to do
                }
            }

            date = new DateTime();
            return false;
        }

        private List<string> GetColumnsToEmailAtRow(Excel._Worksheet sheet, int rowIndex)
        {
            var columns = new List<string>(excelParams.columnsToEmail?.Count ?? 0);
            if (excelParams.columnsToEmail != null)
            {
                foreach (string columnIndex in excelParams.columnsToEmail)
                {
                    try
                    {
                        Excel.Range cell = sheet.Cells[rowIndex, columnIndex];
                        if (cell.Value != null && cell.Value is DateTime)
                        {
                            columns.Add(cell.Value?.ToString("d", CultureInfo.CreateSpecificCulture("ro-RO")) ?? string.Empty);
                        }
                        else
                        {
                            columns.Add(cell.Value?.ToString() ?? string.Empty);
                        }
                    }
                    catch (COMException ex)
                    {
                        logger?.LogE($"Wrong index in accessing cells; Sheet: {sheet.Name}; " +
                            $"Row: {rowIndex}; Column: {columnIndex}; Message: {ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        logger?.LogE($"Exception in GetColumnsToEmailAtRow: {ex.Message}");
                    }
                }
            }

            return columns;
        }

        private List<string> GetHeaderColumnsToEmail()
        {
            var headerRow = new List<string>((excelParams.columnsToEmail?.Count ?? 0) + 1)
            {
                "Sheet & Row"
            };
            headerRow.AddRange(GetColumnsToEmailAtRow(sheets[0], excelParams.headerRow));
            return headerRow;
        }
    }
}
