using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace Saxcel
{
    public class XlsxReader : IDisposable
    {
        readonly SpreadsheetDocument _spreadsheetDocument;
        readonly XlsxReaderConfiguration _configuration;
        readonly WorksheetPart _workSheetPart;
        readonly WorkbookPart _workbookPart;
        readonly Sheet _sheet;

        XlsxReaderStrategy _strategy;
        bool _readHasStarted;

        /// <summary>
        /// Returns the column the reader is currently on.
        /// </summary>
        public string CurrentColumn { get; private set; }

        /// <summary>
        /// Returns the row the reader is currently on.
        /// </summary>
        public int CurrentRow { get; private set; }

        /// <summary>
        /// Creates an instance of the XlsxReader and opens the file in the specified path.
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="sheetname"></param>
        /// <param name="fileMode"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public XlsxReader(string filepath, string sheetname)
        {
            // Check the file extension
            if (!filepath.Split('.').Last().IsOneOf("xlsx", "xls")) throw new InvalidOperationException("Invalid file extension.");

            // Open the workbook and the sheet with the specified name
            _spreadsheetDocument = SpreadsheetDocument.Open(filepath, isEditable: false);
            _workbookPart = _spreadsheetDocument.WorkbookPart;
            _sheet = _workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name.Equals(sheetname)).FirstOrDefault();

            // If sheet not found
            if (_sheet == null) throw new ArgumentException($"Sheet '{sheetname}' does not exist in the workbook.", nameof(sheetname));

            // Load the worksheet part (which contains the actual sheet data)
            _workSheetPart = (WorksheetPart)(_workbookPart.GetPartById(_sheet.Id));
        }

        /// <summary>
        /// Creates an instance of the XlsxReader with specified configurations and opens the file in the specified path.
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="sheetname"></param>
        /// <param name="configure"></param>
        public XlsxReader(string filepath, string sheetname, Action<XlsxReaderConfiguration> configure) : this(filepath, sheetname)
        {          
            configure.Invoke(_configuration);
        }

        /// <summary>
        /// Returns the cell values from the specified column.
        /// </summary>
        /// <param name="column"></param>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        public bool IsReading(string column, out string cellValue)
        {
            // If reader has not yet been started
            if (!_readHasStarted)
            {
                _strategy = new SingleColumnStrategy() { StartColumn = column, EndColumn = column };

                InitializeStrategy();
                Task.Run(() => _strategy.Execute());

                _readHasStarted = true;
            }

            _strategy.HasNewValue = false;
            _strategy.ReadingPaused = false;

            // While the reader is fetching the next cell value
            while (!_strategy.EndOfFileReached &&
                   !_strategy.HasNewValue &&
                   !_strategy.ReadingPaused)
            { } // Wait here

            // Set out cell value, column and row
            cellValue = _strategy.CurrentValue;
            CurrentColumn = _strategy.CurrentColumn;
            CurrentRow = _strategy.CurrentRow;

            // Return true if there are still things left to be read
            return _strategy.EndOfFileReached ? false : true;
        }

        /// <summary>
        /// Returns the cell values in and between the specified columns.
        /// </summary>
        /// <param name="column"></param>
        /// <param name="toColumn"></param>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        public bool IsReading(string column, string toColumn, out string cellValue)
        {
            // If reader has not yet been started
            if (!_readHasStarted)
            {
                _strategy = new MultipleColumnStrategy() { StartColumn = column, EndColumn = toColumn };

                InitializeStrategy();
                Task.Run(() => _strategy.Execute());

                _readHasStarted = true;
            }

            _strategy.HasNewValue = false;
            _strategy.ReadingPaused = false;

            // While the reader is fetching the next cell value
            while (!_strategy.EndOfFileReached &&
                   !_strategy.HasNewValue &&
                   !_strategy.ReadingPaused)
            { } // Wait here

            // Set the current cell value, column and row
            cellValue = _strategy.CurrentValue;
            CurrentColumn = _strategy.CurrentColumn;
            CurrentRow = _strategy.CurrentRow;

            // Return true if there are still things left to be read
            return _strategy.EndOfFileReached ? false : true;
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
            (string from, string to) = range.GetCellRange();
            var startColumn = from.SplitAlphabetsAndNumbers().First();
            var endColumn = to.SplitAlphabetsAndNumbers().First();
            var endColumnLastRowNum = int.Parse(to.SplitAlphabetsAndNumbers().Last());

            // If reader has not yet been started
            if (!_readHasStarted)
            {
                _strategy = new CellRangeStrategy() { StartColumn = startColumn, EndColumn = endColumn, CellRange = range, MaximumRow = endColumnLastRowNum };

                InitializeStrategy();
                Task.Run(() => _strategy.Execute());

                _readHasStarted = true;
            }

            _strategy.HasNewValue = false;
            _strategy.ReadingPaused = false;

            // While the reader is fetching the next cell value
            while (!_strategy.EndOfFileReached && 
                   !_strategy.HasNewValue &&
                   !_strategy.ReadingPaused)
            { } // Wait here

            // Set the current cell value, column and row
            cellValue = _strategy.CurrentValue;
            CurrentColumn = _strategy.CurrentColumn;
            CurrentRow = _strategy.CurrentRow;

            // Return true if there are still things left to be read
            return _strategy.EndOfFileReached ? false : true;
        }

        /// <summary>
        /// Initializes the reader with specified strategy and starts up the read task.
        /// </summary>
        void InitializeStrategy()
        {
            _strategy.WorkbookPart = _workbookPart;
            _strategy.WorksheetPart = _workSheetPart;
            _strategy.DateParser = _configuration?.DateParser ?? new DateParser();
            _strategy.TimeParser = _configuration?.TimeParser ?? new TimeParser();
            _strategy.NumberParser = _configuration?.NumberParser ?? new NumberParser();
            _strategy.StringParser = _configuration?.StringParser ?? new StringParser();

            // Load the Shared String Table into IStringParser's dictionary
            int index = 0;
            foreach (SharedStringItem ssItem in _workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
            {
                _strategy.StringParser.SharedStringTable.Add(index, ssItem.InnerText);
                index++;
            }

            // Get all workbook's custom cell formattings for dates, times, numbers etc
            // and add them to the corresponding parser.
            foreach (var (formatId, formatCode) in _workbookPart.GetCustomCellFormattings())
            {
                if (formatCode.Equals("GENERAL"))
                    _strategy.StringParser.AddFormat((formatId, formatCode));

                else if (formatCode.StartsWith("[$-F800]"))
                    _strategy.DateParser.AddFormat((formatId, formatCode));

                else if (formatCode.StartsWith("[$-F400]"))
                    _strategy.TimeParser.AddFormat((formatId, formatCode));

                else
                    _strategy.NumberParser.AddFormat((formatId, formatCode));
            }
        }

        /// <summary>
        /// Closes the document and releases all resources.
        /// </summary>
        public void Dispose()
        {
            _spreadsheetDocument.Dispose();
        }
    }
}
