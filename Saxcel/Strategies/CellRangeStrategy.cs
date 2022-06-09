using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace Saxcel
{
    internal class CellRangeStrategy : XlsxReaderStrategy
    {
        public CellRangeStrategy(XlsxReader reader) : base(reader)
        { }

        public override async Task BeginRead()
        {
            await Task.Run(() =>
            {
                // Instantiate the OpenXmlReader for reading the worksheet data
                using (Reader = OpenXmlReader.Create(worksheetPart))
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
                            cell = (Cell)Reader.LoadCurrentElement();

                            currentColumn = cell.GetColumnName();

                            if (!currentColumn.Equals(StartColumn, StringComparison.OrdinalIgnoreCase) ||
                                !cell.IsInRange(CellRange))
                                continue;

                            // Set the public column
                            CurrentColumn = currentColumn;

                            // Get cell format for the current cell
                            workbookPart.TryGetCellFormat(cell, out cellFormat);

                            if (StringParser.TryGetValueAndFormat(cell, cellFormat, out ValueTuple<string, string> stringResult))
                            {
                                (string text, string format) = stringResult;
                                CurrentValue = text;
                            }
                            else if (TimeParser.TryGetValueAndFormat(cell, cellFormat, out ValueTuple<DateTime, string> timeResult))
                            {
                                (DateTime time, string format) = timeResult;
                                CurrentValue = time.ToString(format);
                            }
                            else if (DateParser.TryGetValueAndFormat(cell, cellFormat, out ValueTuple<DateTime, string> dateResult))
                            {
                                (DateTime date, string format) = dateResult;
                                CurrentValue = date.ToString(format);
                            }
                            else if (NumberParser.TryGetValueAndFormat(cell, cellFormat, out ValueTuple<decimal, string> numberResult))
                            {
                                (decimal number, string format) = numberResult;
                                CurrentValue = number.ToString(format);
                            }

                            // Set a flag that indicates we have read a new value
                            HasNewValue = true;

                            // Pause reading
                            AllowedToContinue = true;

                            // Pause here until pause flag is set to false
                            while (AllowedToContinue) { };

                        } while (Reader.ReadNextSibling());
                    }
                }
            });

            // End of column reached. Should we read the next column?
            if (!StartColumn.Equals(EndColumn, StringComparison.OrdinalIgnoreCase))
            {
                // Set the next column to be read
                StartColumn = StartColumn.GetNextColumn();

                // Read the column
                if (!EndOfFileReached)
                {
                    BeginRead();
                }
            }
            else
            {
                EndOfFileReached = true;
            }
        }
    }
}
