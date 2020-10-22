using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using StringMatrix = System.Collections.Generic.List<System.Collections.Generic.List<string>>;


namespace ReadSendProject.ExcelReader
{
    class NPOIExcelReader : AbstractExcelReader
    {
        private IWorkbook workbook;
        private readonly List<ISheet> sheets = new List<ISheet>();

        public NPOIExcelReader(ExcelReaderParameters parameters)
            : base(parameters)
        {
        }

        public override StringMatrix Get()
        {
            if (TryOpenExcel())
            {
                GetSheets();
                StringMatrix result = null;

                if (sheets.Count != 0)
                {
                    int estimatedCapacity = 10;
                    result = new StringMatrix(estimatedCapacity)
                    {
                        GetHeaderRowColumnsToEmail()
                    };
                    result.AddRange(ProcessSheets());
                }

                ClearSheets();
                CloseExcel();
                return result;
            }
            return null;
        }

        private StringMatrix ProcessSheets()
        {
            var result = new StringMatrix();

            foreach (var sheet in sheets)
            {
                int startRowIndex = excelParameters.headerRow + 1;
                int endRowIndex = sheet.LastRowNum;

                for (int rowIndex = startRowIndex; rowIndex <= endRowIndex; ++rowIndex)
                {
                    var row = sheet.GetRow(rowIndex);
                    if (row == null)
                    {
                        continue;
                    }
                    if (IsRowExpired(row))
                    {
                        var cells = new List<string>((excelParameters.columnsIndexesToEmail?.Count ?? 0) + 1)
                        {
                            $"{sheet.SheetName} {rowIndex + 1}"
                        };
                        cells.AddRange(GetColumnsToEmailAtRow(sheet, rowIndex));

                        result.Add(cells);
                    }
                }
            }

            return result;
        }

        private bool IsRowExpired(IRow row)
        {
            if (excelParameters.columnsIndexesToCheckDate != null)
            {
                foreach (int columnIndex in excelParameters.columnsIndexesToCheckDate)
                {
                    var cell = row.GetCell(columnIndex);

                    if (TryGetDate(cell, out DateTime cellDate))
                    {
                        try
                        {
                            TimeSpan timeDifference = cellDate - DateTime.Today;
                            if (excelParameters.daysUntilExpirationCheck?.Contains(timeDifference.Days) ?? false)
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

        bool TryGetDate(ICell cell, out DateTime date)
        {
            switch (cell?.CellType ?? CellType.Unknown)
            {
                case CellType.Numeric:
                    try
                    {
                        date = cell.DateCellValue;
                        return true;
                    }
                    catch
                    {
                        try
                        {
                            date = DateTime.FromOADate(cell.NumericCellValue); // maybe??
                            return true;
                        }
                        catch (ArgumentException)
                        {
                            // Nothing to do
                        }
                    }
                    break;
                case CellType.String:
                    return DateTime.TryParseExact(
                        cell.StringCellValue,
                        excelParameters.dateFormats.ToArray(),
                        null,
                        DateTimeStyles.None,
                        out date
                    );
                case CellType.Formula:
                case CellType.Blank:
                case CellType.Boolean:
                case CellType.Error:
                case CellType.Unknown:
                default:
                    break;
            }

            date = new DateTime();
            return false;
        }

        private bool TryOpenExcel()
        {
            try
            {
                using (var file = new FileStream(excelParameters.path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    workbook = WorkbookFactory.Create(file, bReadonly: true);
                }
                return true;
            }
            catch (Exception ex)
            {
                logger?.LogE($"Exception while trying to open excel; Message: {ex.Message}");
                return false;
            }
        }

        private void CloseExcel() => workbook?.Close();

        private void GetSheets()
        {
            if (excelParameters.sheetsNames.IsNullOrEmpty() == false)
            {
                foreach (string sheetName in excelParameters.sheetsNames)
                {
                    try
                    {
                        sheets.Add(workbook.GetSheet(sheetName));
                    }
                    catch (Exception)
                    {
                        logger?.LogE($"Incorrect name given for a worksheet: {sheetName}");
                        continue;
                    }
                }
            }

            if (excelParameters.sheetsIndexes.IsNullOrEmpty() == false)
            {
                foreach (int index in excelParameters.sheetsIndexes)
                {
                    try
                    {
                        sheets.Add(workbook.GetSheetAt(index));
                    }
                    catch (Exception)
                    {
                        logger?.LogE($"Incorrect index given for a worksheet: {index}");
                        continue;
                    }
                }
            }

            if (sheets.Count == 0)
            {
                logger?.LogI("No sheets to process.");
            }
        }

        private void ClearSheets() => sheets.Clear();

        List<string> GetColumnsToEmailAtRow(ISheet sheet, int rowIndex)
        {
            var columns = new List<string>(excelParameters.columnsIndexesToEmail?.Count ?? 0);

            if (excelParameters.columnsIndexesToEmail == null
                || sheet == null
                || sheet.GetRow(rowIndex) == null)
            {
                return columns;
            }

            try
            {
                var row = sheet.GetRow(rowIndex);

                foreach (int colIndex in excelParameters.columnsIndexesToEmail)
                {
                    var cell = row.GetCell(colIndex);
                    switch (cell?.CellType ?? CellType.Unknown)
                    {
                        case CellType.Numeric:
                            if (excelParameters.columnsIndexesToCheckDate?.Contains(colIndex) ?? false)
                            {
                                try
                                {
                                    columns.Add(cell.DateCellValue.ToString("d", CultureInfo.CreateSpecificCulture("ro-RO")));
                                }
                                catch (NullReferenceException)
                                {
                                    try
                                    {
                                        var date = DateTime.FromOADate(cell.NumericCellValue);
                                        columns.Add(date.ToString("d", CultureInfo.CreateSpecificCulture("ro-RO")));
                                    }
                                    catch (ArgumentException)
                                    {
                                        columns.Add(cell.NumericCellValue.ToString());
                                    }
                                }
                            }
                            else
                            {
                                columns.Add(cell.NumericCellValue.ToString());
                            }
                            break;
                        case CellType.String:
                            columns.Add(cell.StringCellValue);
                            break;
                        case CellType.Boolean:
                        case CellType.Blank:
                        case CellType.Formula:
                        case CellType.Unknown:
                        case CellType.Error:
                        default:
                            columns.Add(string.Empty);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                logger?.LogE($"Exception in GetColumnsToEmailAtRow; Message {ex.Message}");
            }

            return columns;
        }

        List<string> GetHeaderRowColumnsToEmail()
        {
            var headerRow = new List<string>((excelParameters.columnsIndexesToEmail?.Count ?? 0) + 1)
            {
                "Sheet & Row"
            };
            headerRow.AddRange(GetColumnsToEmailAtRow(sheets[0], excelParameters.headerRow));
            return headerRow;
        }
    }
}
