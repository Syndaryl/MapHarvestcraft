using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Spreadsheet.Worksheet;
using S = DocumentFormat.OpenXml.Spreadsheet.Sheet;
using E = DocumentFormat.OpenXml.OpenXmlElement;
using A = DocumentFormat.OpenXml.OpenXmlAttribute;
using System.Text.RegularExpressions;
using Extensions;

namespace ReadExcelFile
{

    public class WorkSheetHandler : IDisposable
    {

        public WorkSheetHandler()
        {
        }
        SpreadsheetDocument excelDocument;

        public SpreadsheetDocument ExcelDocument
        {
            get
            {
                return excelDocument;
            }

            private set
            {
                if (excelDocument == null)
                {
                    excelDocument = value;
                }
                else if (excelDocument.Equals(value)) return;
                else
                {
                    excelDocument.Dispose();
                    excelDocument = value;
                }
            }
        }

        public EventHandler On100thRow { get; internal set; }

        public void SetExcelDocument(string filename)
        {
            // Create instance of OpenSettings
            OpenSettings openSettings = new OpenSettings
            {

                // Add the MarkupCompatibilityProcessSettings
                MarkupCompatibilityProcessSettings =
                new MarkupCompatibilityProcessSettings(
                    MarkupCompatibilityProcessMode.ProcessAllParts,
                    FileFormatVersions.Office2013)
            };

            ExcelDocument =
                SpreadsheetDocument.Open(filename, false, openSettings);
        }

        public List<List<string>> Sheet(string sheetName, bool IgnoreEmptyRows)
        {
            GetSheetPart(sheetName,  out SharedStringTablePart stringTable, out WorksheetPart wsPart);
            W worksheet = wsPart.Worksheet;

            IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Elements<Row>();
            uint firstRowNum = rows.Select(r => r.RowIndex).Min();
            uint lastRowNum = rows.Select(r => r.RowIndex).Max();
            var columns = rows.SelectMany(r => r.Elements<Cell>()).Distinct().OrderBy(c => ColumnComparer.GetColumnName(c.CellReference), new ColumnComparer()).Select(c => ColumnComparer.GetColumnName(c.CellReference));

            string firstColumn = columns.First();
            string lastColumn = columns.Last();
            return SheetRange(sheetName, firstRowNum, lastRowNum, firstColumn, lastColumn, IgnoreEmptyRows, stringTable, wsPart);
        }

        public List<List<string>> SheetRange(string sheetName, string firstCellName, string lastCellName, bool IgnoreEmptyRows = true)
        {

            uint firstRowNum = ColumnComparer.GetRowIndex(firstCellName);
            uint lastRowNum = ColumnComparer.GetRowIndex(lastCellName);
            string firstColumn = ColumnComparer.GetColumnName(firstCellName);
            string lastColumn = ColumnComparer.GetColumnName(lastCellName);
            return SheetRange(sheetName, firstRowNum, lastRowNum, firstColumn, lastColumn, IgnoreEmptyRows);
        }

        private List<List<string>> SheetRange(string sheetName, uint firstRowNum, uint lastRowNum, string firstColumn, string lastColumn, bool IgnoreEmptyRows, SharedStringTablePart sharedStringTablePart=null, WorksheetPart worksheetPart=null)
        {
            SharedStringTablePart stringTable;
            WorksheetPart wsPart;
            if (sharedStringTablePart == null || worksheetPart == null) { 
                GetSheetPart(sheetName, out stringTable, out wsPart);
            } else
            {
                stringTable = sharedStringTablePart;
                wsPart = worksheetPart;
            }

            W worksheet = wsPart.Worksheet;

            uint rowCount = lastRowNum - firstRowNum;
            var lastColNum = GetExcelColumnNumber(lastColumn);
            var firstColNum = GetExcelColumnNumber(firstColumn);
            uint colCount = lastColNum - firstColNum;

            var result = new List<List<string>>((int)rowCount);

            // Iterate through the cells within the range and do whatever.
            foreach (int row in Enumerable.Range((int)firstRowNum - 1, (int)rowCount))
            {
                var newRow = new List<string>((int)colCount);
                if (result.Count % 100 == 0)
                {
                    On100thRow.Invoke(this, new EventArgs());
                }
                foreach (var columnID in Enumerable.Range((int)firstColNum - 1, (int)colCount).Select(e => GetExcelColumnName(e)))
                {
                    var cellRef = String.Format("{0}{1}", columnID, row);
                    var cell = wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == cellRef).FirstOrDefault();
                    newRow.Add(DecodedCell(cell, stringTable)?.Trim());
                }
                if (
                    (!IgnoreEmptyRows) || (!newRow.All(c => string.IsNullOrEmpty(c) || c.Equals("undef")))
                    )

                {
                    result.Add(newRow);
                }
            }
            return result;
        }

        private void GetSheetPart(string sheetName, out SharedStringTablePart stringTable, out WorksheetPart wsPart)
        {
            var wbPart = excelDocument.WorkbookPart;
            Sheets sheets = wbPart.Workbook.Sheets;
            S sheet = wbPart.Workbook.Descendants<S>().Where(s => s.Name == sheetName).First();
            // Throw an exception if there is no sheet.
            if (sheet == null)
            {
                throw new ArgumentException("sheetName is invalid for this workbook");
            }
            stringTable = wbPart.GetPartsOfType<SharedStringTablePart>()
                .FirstOrDefault();

            // Retrieve a reference to the worksheet part.
            wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
        }

        public void DumpSheetRange<T>(T output, string filename, string sheetName, string firstCellName, string lastCellName) where T : TextWriter
        {

            // Create instance of OpenSettings
            OpenSettings openSettings = new OpenSettings
            {

                // Add the MarkupCompatibilityProcessSettings
                MarkupCompatibilityProcessSettings =
                new MarkupCompatibilityProcessSettings(
                    MarkupCompatibilityProcessMode.ProcessAllParts,
                    FileFormatVersions.Office2013)
            };

            // Open the document with OpenSettings
            using (SpreadsheetDocument excelDocument =
                SpreadsheetDocument.Open(filename,
                    false,
                    openSettings))
            {
                var wbPart = excelDocument.WorkbookPart;
                Sheets sheets = wbPart.Workbook.Sheets;
                foreach (E sheetElem in sheets)
                {
                    foreach (A attr in sheetElem.GetAttributes())
                    {
                        output.Write("{0}: {1}\t", attr.LocalName, attr.Value);
                    }
                    output.WriteLine();
                }
                S sheet = wbPart.Workbook.Descendants<S>().Where(s => s.Name == sheetName).First();
                // Throw an exception if there is no sheet.
                if (sheet == null)
                {
                    throw new ArgumentException("sheetName is invalid for this workbook");
                }
                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart =
                    (WorksheetPart)wbPart.GetPartById(sheet.Id);

                W worksheet = wsPart.Worksheet;
                var stringTable =
                    wbPart.GetPartsOfType<SharedStringTablePart>()
                    .FirstOrDefault();

                uint firstRowNum = ColumnComparer.GetRowIndex(firstCellName);
                uint lastRowNum = ColumnComparer.GetRowIndex(lastCellName);
                string firstColumn = ColumnComparer.GetColumnName(firstCellName);
                string lastColumn = ColumnComparer.GetColumnName(lastCellName);

                string headers = firstColumn;
                while (ColumnComparer.CompareColumn(headers, lastColumn) <= 0)
                {
                    output.Write("{0}\t", headers);
                    headers = ColumnComparer.NextColumn(headers);
                }
                output.WriteLine();

                //// Iterate through the cells within the range and do whatever.
                foreach (int row in Enumerable.Range((int)firstRowNum, (int)lastRowNum))
                {
                    string col = firstColumn;
                    while (ColumnComparer.IsColumnInRange(firstColumn, lastColumn, col))
                    {
                        var cellRef = String.Format("{0}{1}", col, row);
                        var cell = wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == cellRef).FirstOrDefault();
                        output.Write("{0}\t", DecodedCell(cell, stringTable));
                        col = ColumnComparer.NextColumn(col);
                    }
                    output.WriteLine();
                }
            }
        }

        private string DecodedCell(Cell cell, SharedStringTablePart stringTable)
        {
            if (cell == null) return "";

            var value = cell?.CellValue?.InnerText;
            if (cell?.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.Boolean:
                        switch (value)
                        {
                            case "0":
                                return "FALSE";
                            default:
                                return "TRUE";
                        }
                    case CellValues.Number:
                        return value;
                    case CellValues.Error:
                        return value;
                    case CellValues.SharedString:
                        if (stringTable != null)
                        {
                            return
                                stringTable.SharedStringTable
                                .ElementAt(int.Parse(value)).InnerText;
                        }
                        else
                        {
                            return value;
                        }
                    case CellValues.String:
                        return value;
                    case CellValues.InlineString:
                        return value;
                    case CellValues.Date:
                        break;
                    default:
                        return "undef";
                }

            }
            return value;
        }


        private string GetExcelColumnName(int Index)
        {
            string range = "";
            if (Index < 0) return range;
            for (int i = 1; Index + i > 0; i = 0)
            {
                range = ((char)(65 + Index % 26)).ToString() + range;
                Index /= 26;
            }
            if (range.Length > 1) range = ((char)((int)range[0] - 1)).ToString() + range.Substring(1);
            return range;
        }

        private uint GetExcelColumnNumber(string columnName)
        {
            //byte[] indexes = new byte[columnName.Length];
            uint value = 0;
            for (int i = columnName.Length-1; i >= 0 && i < columnName.Length; i--)
            {
                byte posVal = (byte)(LetterToNumber(columnName[i])+ 1);
                value += (uint)Math.Pow(posVal, columnName.Length - i);
            }

            return value;
        }

        /// <summary>
        /// Returns zero-indexed number indicating letter of the alphabet.
        /// </summary>
        /// <param name="letter">The letter to check as a <see cref="char"/></param>
        /// <returns></returns>
        private byte LetterToNumber(char letter)
        {
            return (byte)(char.ToUpper(letter) - 'A' );
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    excelDocument.Dispose();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.

                // TODO: set large fields to null.


                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~WorkSheetHandler() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion
    }


}