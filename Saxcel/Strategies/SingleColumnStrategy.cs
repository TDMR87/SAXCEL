using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace Saxcel
{
    /// <summary>
    /// This strategy reads the values of one single column.
    /// </summary>
    internal class SingleColumnStrategy : XlsxReaderStrategy
    {
        public SingleColumnStrategy(XlsxReader reader) : base(reader) 
        { }

        public override async Task BeginRead()
        {
            await Task.Run(() =>
            {
                // Open the file
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

                            do // Read all sibling elements of the Row
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

                                // Pause reading
                                AllowedToContinue = false;
                                
                                // Pause here until the pause flag is set to false
                                while (AllowedToContinue == false) { };

                            } while (Reader.ReadNextSibling());
                        }
                    }
                }
            });

            // End of file reached
            EndOfFileReached = true;
        }
    }
}
