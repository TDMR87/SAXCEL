using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Saxcel
{
    /// <summary>
    /// The XlsxReader class opens a .xslx file for reading and 
    /// iterates through the cells yielding the value of each cell.
    /// </summary>
    public class XlsxReader : IDisposable
    {
        internal readonly SpreadsheetDocument spreadsheetDocument;
        internal readonly XlsxReaderConfiguration configuration;
        internal readonly WorksheetPart workSheetPart;
        internal readonly WorkbookPart workbookPart;
        internal readonly Sheet sheet;

        private XlsxReaderStrategy _readingStrategy;
        private bool _readingStarted;

        /// <summary>
        /// Returns the column the reader is currently on.
        /// </summary>
        public string CurrentColumn { get; private set; }

        /// <summary>
        /// Returns the row the reader is currently on.
        /// </summary>
        public int CurrentRow { get; private set; }

        /// <summary>
        /// Creates an instance of the XlsxReader.
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="sheetname"></param>
        /// <param name="fileMode"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public XlsxReader(string filepath, string sheetname)
        {
            // Check the file extension
            if (!filepath.Split('.').Last().IsOneOf("xlsx")) 
                throw new InvalidOperationException("Invalid file extension.");

            // Open the workbook and the sheet
            spreadsheetDocument = SpreadsheetDocument.Open(filepath, isEditable: false);
            workbookPart = spreadsheetDocument.WorkbookPart;
            sheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name.Equals(sheetname)).FirstOrDefault();

            // If sheet not found
            if (sheet == null) throw new ArgumentException($"Sheet '{sheetname}' does not exist in the workbook.", nameof(sheetname));

            // Load the worksheet part object (which contains the actual sheet data)
            workSheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
        }

        /// <summary>
        /// Creates an instance of the XlsxReader with specified configurations.
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="sheetname"></param>
        /// <param name="configure"></param>
        public XlsxReader(string filepath, string sheetname, Action<XlsxReaderConfiguration> configure) : this(filepath, sheetname)
        {
            configure.Invoke(configuration);
        }

        /// <summary>
        /// While true, yields the cell values one by one from the specified column.
        /// </summary>
        /// <param name="column"></param>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        public bool IsReading(string column, out string cellValue)
        {
            if (column.Any(c => !Char.IsLetter(c)))
            {
                throw new ArgumentException("Column name cannot contain any special characters or numbers.", nameof(column));
            }

            // If reader has not yet been started, start it.
            if (_readingStrategy == null)
            {
                _readingStrategy = new SingleColumnStrategy(this) { StartColumn = column, EndColumn = column };
                _ = _readingStrategy.BeginRead();
            }

            _readingStrategy.AllowedToContinue = true;

            // Wait while the reader is fetching the next value
            while (_readingStrategy.AllowedToContinue) { }

            // Set the cell value, current column and the current row
            cellValue = _readingStrategy.CurrentValue;
            CurrentColumn = _readingStrategy.CurrentColumn;
            CurrentRow = _readingStrategy.CurrentRow;

            // Return false if end of the file has been reached,
            // return true if there are still values left to be read
            return _readingStrategy.EndOfFileReached ? false : true;
        }

        /// <summary>
        /// Returns the cell values from the specified columns and all the columns
        /// between them.
        /// </summary>
        /// <param name="fromColumn"></param>
        /// <param name="toColumn"></param>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        public bool IsReading(string fromColumn, string toColumn, out string cellValue)
        {
            if (fromColumn.Any(c => !Char.IsLetter(c)) || toColumn.Any(c => !Char.IsLetter(c)))
            {
                throw new ArgumentException("Column names cannot contain any special characters or numbers.");
            }

            // If reader has not yet been started, start it
            if (_readingStrategy == null)
            {
                _readingStrategy = new MultipleColumnStrategy(this) { StartColumn = fromColumn, EndColumn = toColumn };
                _ = _readingStrategy.BeginRead();
            }

            _readingStrategy.AllowedToContinue = true;

            while (_readingStrategy.AllowedToContinue){}

            // Set the current cell value, current column and current row
            cellValue = _readingStrategy.CurrentValue;
            CurrentColumn = _readingStrategy.CurrentColumn;
            CurrentRow = _readingStrategy.CurrentRow;

            // Return false if end of the file has been reached,
            // return true if there are still values left to be read
            return _readingStrategy.EndOfFileReached ? false : true;
        }

        /// <summary>
        /// Returns the cell values inside the specified range of columns and rows.
        /// The range must be specified in a format that specifies the column and 
        /// row number. e.g. A10:C99.
        /// </summary>
        /// <param name="range"></param>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        public bool IsReadingRange(string range, out string cellValue)
        {
            (string fromCell, string toCell) = range.GetCellRange();
            var startColumn = fromCell.SplitAlphabetsAndNumbers().First();
            var endColumn = toCell.SplitAlphabetsAndNumbers().First();
            var endColumnLastRowNum = int.Parse(toCell.SplitAlphabetsAndNumbers().Last());

            // If reader has not yet been started
            if (_readingStrategy == null)
            {
                _readingStrategy = new CellRangeStrategy(this) 
                { 
                    StartColumn = startColumn, 
                    EndColumn = endColumn, 
                    CellRange = range, 
                    MaximumRow = endColumnLastRowNum 
                };

                _ = _readingStrategy.BeginRead();
            }

            _readingStrategy.AllowedToContinue = true;

            // Wait while the reader is fetching the next value
            while (_readingStrategy.AllowedToContinue) { }

            // Set the current cell value, current column and current row
            cellValue = _readingStrategy.CurrentValue;
            CurrentColumn = _readingStrategy.CurrentColumn;
            CurrentRow = _readingStrategy.CurrentRow;

            // Return false if end of the file has been reached,
            // return true if there are still values left to be read
            return _readingStrategy.EndOfFileReached ? false : true;
        }

        /// <summary>
        /// Closes the XlsxReader and the loaded .xlsx document and releases all resources.
        /// </summary>
        public void Dispose()
        {
            spreadsheetDocument.Dispose();
        }
    }
}
