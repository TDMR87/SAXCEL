using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;

namespace Saxcel
{
    internal class CellRangeStrategy : XlsxReaderStrategy
    {
        public CellRangeStrategy(WorkbookPart workbookPart, WorksheetPart worksheetPart, XlsxReaderConfiguration configuration) : 
            base(workbookPart, worksheetPart, configuration)
        { }

        public override void Execute()
        {
            // Instantiate the OpenXmlReader for reading the worksheet data
            using (Reader = OpenXmlReader.Create(WorksheetPart))
            {
                // Read the xlsx file from beginning to the end
                while (Reader.Read())
                {
                    // Skip other than Row elements
                    if (Reader.ElementType != typeof(Row)) continue;

                    // Set current row
                    CurrentRow = int.Parse(Reader.Attributes.First(attr => attr.LocalName == "r").Value);

                    // If we reached the end of the range
                    if (StartColumn == EndColumn && CurrentRow > MaximumRow)
                    {
                        // Stop reading
                        break;
                    }

                    // Read a Row element
                    Reader.ReadFirstChild();

                    do // Read all siblings of the Row (e.g. cells)
                    {
                        // Skip if the element is not a cell
                        if (Reader.ElementType != typeof(Cell)) continue;

                        // Load the cell element
                        _cell = (Cell)Reader.LoadCurrentElement();

                        _currentColumn = _cell.GetColumnName();

                        if (!_currentColumn.Equals(StartColumn, StringComparison.OrdinalIgnoreCase) || 
                            !_cell.IsInRange(CellRange)) 
                                continue;

                        // Set the public column
                        CurrentColumn = _currentColumn;

                        // Get cell format for the current cell
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

                    } while (Reader.ReadNextSibling());
                }
            }

            // End of column reached. Should we read the next column?
            if (!StartColumn.Equals(EndColumn, StringComparison.OrdinalIgnoreCase))
            {
                // Get the next column to be read
                StartColumn = StartColumn.GetNextColumn();

                // Read the column
                if (!EndOfFileReached)
                {
                    Execute();
                }
            }
            else
            {
                EndOfFileReached = true;
            }
        }
    }
}
