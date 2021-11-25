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

        public WorkbookPart WorkbookPart { get; set; }
        public WorksheetPart WorksheetPart { get; set; }
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
        public bool ReadingPaused { get; set; }

        /// <summary>
        /// Executes the strategy.
        /// </summary>
        public abstract void Execute();
    }
}
