using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace Saxcel
{
    /// <summary>
    /// Base class for all reading strategies.
    /// </summary>
    internal abstract class XlsxReaderStrategy
    {
        protected CellFormat _cellFormat;
        protected Cell _cell;
        protected string _currentColumn;
        protected XlsxReaderConfiguration _configuration;

        public XlsxReaderStrategy(WorkbookPart workbookPart, WorksheetPart worksheetPart, XlsxReaderConfiguration configuration)
        {
            WorkbookPart = workbookPart;
            WorksheetPart = worksheetPart;
            _configuration = configuration;

            Initialize();
        }

        protected WorkbookPart WorkbookPart { get; set; }
        protected WorksheetPart WorksheetPart { get; set; }
        public OpenXmlReader Reader { get; set; }
        public IStringParser StringParser { get; set; }
        public IParser<decimal> NumberParser { get; set; }
        public IParser<DateTime> TimeParser { get; set; }
        public IParser<DateTime> DateParser  { get; set; }
        public int CurrentRow { get; set; }
        public int MaximumRow { get; set; }
        public string EndColumn { get; set; }
        public string StartColumn { get; set; }
        public string CellRange { get; set; }
        public string CurrentValue { get; protected set; }
        public string CurrentColumn { get; protected set; }
        public bool EndOfFileReached { get; set; }
        public bool HasNewValue { get; set; }
        public bool OnPause { get; set; }

        /// <summary>
        /// Executes the strategy.
        /// </summary>
        public abstract void Execute();

        private void Initialize()
        {
            DateParser = _configuration?.DateParser ?? new DateParser();
            TimeParser = _configuration?.TimeParser ?? new TimeParser();
            NumberParser = _configuration?.NumberParser ?? new NumberParser();
            StringParser = _configuration?.StringParser ?? new StringParser();

            // Load the xlsx file's Shared String Table into the reading strategy's string parser
            int index = 0;
            foreach (SharedStringItem ssItem in WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
            {
                StringParser.SharedStringTable.Add(index, ssItem.InnerText);
                index++;
            }

            // Get all the custom cell formattings in the loaded xlsx file (dates, times, numbers etc)
            // and adds them to the corresponding parsers.
            foreach (var (formatId, formatCode) in WorkbookPart.GetCustomCellFormattings())
            {
                if (formatCode.Equals("GENERAL"))
                    StringParser.AddFormat((formatId, formatCode));

                else if (formatCode.StartsWith("[$-F800]"))
                    DateParser.AddFormat((formatId, formatCode));

                else if (formatCode.StartsWith("[$-F400]"))
                    TimeParser.AddFormat((formatId, formatCode));

                else
                    NumberParser.AddFormat((formatId, formatCode));
            }
        }
    }
}
