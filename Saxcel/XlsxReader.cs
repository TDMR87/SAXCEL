using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Saxcel
{
    /// <summary>
    /// The XlsxReader class is responsible for iterating through the cells
    /// in a .xlsx file and yielding the value of each cell.
    /// 
    /// 
    /// </summary>
    public class XlsxReader : IDisposable
    {
        readonly SpreadsheetDocument _spreadsheetDocument;
        readonly XlsxReaderConfiguration _configuration;
        readonly WorksheetPart _workSheetPart;
        readonly WorkbookPart _workbookPart;
        readonly Sheet _sheet;

        XlsxReaderStrategy _readerStrategy;
        bool _readStarted;

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
            _spreadsheetDocument = SpreadsheetDocument.Open(filepath, isEditable: false);
            _workbookPart = _spreadsheetDocument.WorkbookPart;
            _sheet = _workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name.Equals(sheetname)).FirstOrDefault();

            // If sheet not found
            if (_sheet == null) throw new ArgumentException($"Sheet '{sheetname}' does not exist in the workbook.", nameof(sheetname));

            // Load the worksheet part object (which contains the actual sheet data)
            _workSheetPart = (WorksheetPart)(_workbookPart.GetPartById(_sheet.Id));
        }

        /// <summary>
        /// Creates an instance of the XlsxReader with specified configurations.
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="sheetname"></param>
        /// <param name="configure"></param>
        public XlsxReader(string filepath, string sheetname, Action<XlsxReaderConfiguration> configure) : this(filepath, sheetname)
        {
            configure.Invoke(_configuration);
        }

        /// <summary>
        /// While true, yields the cell values from the specified column.
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
            if (!_readStarted)
            {
                _readerStrategy = new SingleColumnStrategy(_workbookPart, _workSheetPart, _configuration) 
                { 
                    StartColumn = column, 
                    EndColumn = column
                };

                Task.Run(() => _readerStrategy.Execute());

                _readStarted = true;
            }

            // Unpause the strategy and wait for results
            _readerStrategy.OnPause = false;

            // While the reader is fetching the next cell value
            while (!_readerStrategy.EndOfFileReached && !_readerStrategy.HasNewValue)
            { } // Wait here

            // Set the cell value, current column and the current row
            cellValue = _readerStrategy.CurrentValue;
            CurrentColumn = _readerStrategy.CurrentColumn;
            CurrentRow = _readerStrategy.CurrentRow;

            // Return false if end of the file has been reached,
            // return true if there are still values left to be read
            return _readerStrategy.EndOfFileReached ? false : true;
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
            if (!_readStarted)
            {
                _readerStrategy = new MultipleColumnStrategy(_workbookPart, _workSheetPart, _configuration) 
                { 
                    StartColumn = fromColumn, 
                    EndColumn = toColumn 
                };

                Task.Run(() => _readerStrategy.Execute());

                _readStarted = true;
            }

            _readerStrategy.HasNewValue = false;
            _readerStrategy.OnPause = false;

            // While the reader is fetching the next cell value
            while (!_readerStrategy.EndOfFileReached &&
                   !_readerStrategy.HasNewValue &&
                   !_readerStrategy.OnPause)
            { } // Wait here

            // Set the current cell value, current column and current row
            cellValue = _readerStrategy.CurrentValue;
            CurrentColumn = _readerStrategy.CurrentColumn;
            CurrentRow = _readerStrategy.CurrentRow;

            // Return false if end of the file has been reached,
            // return true if there are still values left to be read
            return _readerStrategy.EndOfFileReached ? false : true;
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
            if (!_readStarted)
            {
                _readerStrategy = new CellRangeStrategy(_workbookPart, _workSheetPart, _configuration) 
                { 
                    StartColumn = startColumn, 
                    EndColumn = endColumn, 
                    CellRange = range, 
                    MaximumRow = endColumnLastRowNum 
                };

                Task.Run(() => _readerStrategy.Execute());

                _readStarted = true;
            }

            _readerStrategy.HasNewValue = false;
            _readerStrategy.OnPause = false;

            // While the reader is fetching the next cell value
            while (!_readerStrategy.EndOfFileReached &&
                   !_readerStrategy.HasNewValue &&
                   !_readerStrategy.OnPause)
            { } // Wait here

            // Set the current cell value, current column and current row
            cellValue = _readerStrategy.CurrentValue;
            CurrentColumn = _readerStrategy.CurrentColumn;
            CurrentRow = _readerStrategy.CurrentRow;

            // Return false if end of the file has been reached,
            // return true if there are still values left to be read
            return _readerStrategy.EndOfFileReached ? false : true;
        }

        /// <summary>
        /// Closes the XlsxReader and the loaded .xlsx document and releases all resources.
        /// </summary>
        public void Dispose()
        {
            _spreadsheetDocument.Dispose();
        }
    }
}
