using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace Saxcel
{
    internal class MultipleColumnStrategy : XlsxReaderStrategy
    {
        public MultipleColumnStrategy(XlsxReader reader) : base(reader)
        { }

        public override async Task BeginRead()
        {
            await Task.Run(() =>
            {
                // Open the file using a strategy
                using (Reader = OpenXmlReader.Create(worksheetPart))
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
                                cell = (Cell)Reader.LoadCurrentElement();

                                // Skip if cell is not in the specified column
                                if (!cell.GetColumnName().Equals(StartColumn, StringComparison.OrdinalIgnoreCase)) continue;

                                // Set current column
                                CurrentColumn = cell.GetColumnName();

                                // Get cell format
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

                                // Set flag to false after pausing and continue reading the next cell
                                HasNewValue = false;

                            } while (Reader.ReadNextSibling());
                        }
                    }
                }
            });

            // Force garbage collection
            GC.Collect();

            // End of column reached. Should we read the next column?
            if (!StartColumn.Equals(EndColumn, StringComparison.OrdinalIgnoreCase))
            {
                // Get the next column to be read
                StartColumn = StartColumn.GetNextColumn();

                // Read the column
                BeginRead();
            }
            else
            {
                EndOfFileReached = true;
            }
        }
    }
}
