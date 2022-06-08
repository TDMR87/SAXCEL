using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace Saxcel
{
    internal class MultipleColumnStrategy : XlsxReaderStrategy
    {
        public MultipleColumnStrategy(WorkbookPart workbookPart, WorksheetPart worksheetPart, XlsxReaderConfiguration configuration) : 
            base(workbookPart, worksheetPart, configuration)
        { }

        public override void Execute()
        {
            // Open the file using a strategy
            using (Reader = OpenXmlReader.Create(WorksheetPart))
            {
                // Read element by element
                while (Reader.Read())
                {
                    // Skip other than Row elements
                    if (Reader.ElementType == typeof(Row))
                    {
                        // Set current row
                        CurrentRow = int.Parse(Reader.Attributes.First(attr => attr.LocalName == "r").Value);

                        // Read a Row element
                        Reader.ReadFirstChild();

                        do // Read all siblings of the Row (e.g. cells)
                        {
                            // Skip if the element is not a cell
                            if (Reader.ElementType != typeof(Cell)) continue;

                            // Load the cell element
                            _cell = (Cell)Reader.LoadCurrentElement();

                            // Skip if cell is not in the specified column
                            if (!_cell.GetColumnName().Equals(StartColumn, StringComparison.OrdinalIgnoreCase)) continue;

                            // Set current column
                            CurrentColumn = _cell.GetColumnName();

                            // Get cell format
                            WorkbookPart.TryGetCellFormat(_cell, out _cellFormat);

                            if (StringParser.TryGetValueAndFormat(_cell, _cellFormat, out ValueTuple<string, string> stringResult))
                            {
                                (string text, string format) = stringResult;
                                CurrentValue = text;
                            }
                            else if (TimeParser.TryGetValueAndFormat(_cell, _cellFormat, out ValueTuple<DateTime, string> timeResult))
                            {
                                (DateTime time, string format) = timeResult;
                                CurrentValue = time.ToString(format);
                            }
                            else if (DateParser.TryGetValueAndFormat(_cell, _cellFormat, out ValueTuple<DateTime, string> dateResult))
                            {
                                (DateTime date, string format) = dateResult;
                                CurrentValue = date.ToString(format);
                            }
                            else if (NumberParser.TryGetValueAndFormat(_cell, _cellFormat, out ValueTuple<decimal, string> numberResult))
                            {
                                (decimal number, string format) = numberResult;
                                CurrentValue = number.ToString(format);
                            }

                            // Set a flag that indicates we have read a new value
                            HasNewValue = true;

                            // Pause reading
                            OnPause = true;

                            // Pause here until pause flag is set to false
                            while (OnPause) { };

                            // Set flag to false after pausing and continue reading the next cell
                            HasNewValue = false;

                        } while (Reader.ReadNextSibling());
                    }
                }
            }

            // Force garbage collection
            GC.Collect();

            // End of column reached. Should we read the next column?
            if (!StartColumn.Equals(EndColumn, StringComparison.OrdinalIgnoreCase))
            {
                // Get the next column to be read
                StartColumn = StartColumn.GetNextColumn();

                // Read the column
                Execute();
            }
            else
            {
                EndOfFileReached = true;
            }
        }
    }
}
