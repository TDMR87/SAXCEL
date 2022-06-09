using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Threading.Tasks;

namespace Saxcel
{
    /// <summary>
    /// Base class for all reading strategies.
    /// </summary>
    internal abstract class XlsxReaderStrategy
    {
        protected XlsxReaderConfiguration readerConfiguration;
        protected WorkbookPart workbookPart;
        protected WorksheetPart worksheetPart;
        protected CellFormat cellFormat;
        protected Cell cell;
        protected string currentColumn;

        public XlsxReaderStrategy(XlsxReader reader)
        {
            workbookPart = reader.workbookPart;
            worksheetPart = reader.workSheetPart;
            readerConfiguration = reader.configuration;

            DateParser = reader.configuration?.DateParser ?? new DateParser();
            TimeParser = reader.configuration?.TimeParser ?? new TimeParser();
            NumberParser = reader.configuration?.NumberParser ?? new NumberParser();
            StringParser = reader.configuration?.StringParser ?? new StringParser();

            Initialize();
        }
        
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
        public bool AllowedToContinue { get; set; }

        public Func<bool> OnNewValueAvailable { get; set; }

        /// <summary>
        /// Begins reading a file using this strategy.
        /// </summary>
        public abstract Task BeginRead();

        /// <summary>
        /// Initializes the strategy with either the default parsers or some user defined parsers.
        /// the some 
        /// </summary>
        private void Initialize()
        {
            // Load the workbook's Shared String Table into the string parser
            int index = 0;
            foreach (SharedStringItem ssItem in workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
            {
                StringParser.SharedStringTable.Add(index, ssItem.InnerText);
                index++;
            }

            // Get all the custom cell formattings in the loaded xlsx file (dates, times, numbers etc)
            // and adds them to the corresponding parsers.
            foreach (var (formatId, formatCode) in workbookPart.GetCustomCellFormattings())
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
