using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Saxcel
{
    internal class StringParser : IStringParser
    {
        /// <summary>
        /// Contains text formats for displaying text values as they were in the .xlsx file.
        /// </summary>
        public Dictionary<int, string> Formats { get; set; } = new Dictionary<int, string>();

        /// <summary>
        /// The shared string table is a key-value dictionary containing
        /// the text values of an .xlsx file.
        /// </summary>
        public Dictionary<int, string> SharedStringTable { get; set; } = new Dictionary<int, string>();

        /// <summary>
        /// Adds a format to the collection of formats.
        /// </summary>
        /// <param name="customFormat"></param>
        public void AddFormat((int key, string value) customFormat)
        {
            Formats.Add(customFormat.key, customFormat.value);
        }

        /// <summary>
        /// Checks that the cell value is text format and returns a ValueTuple<string, string>
        /// that contains the value itself (string type) and the formatting for displaying that
        /// value. 
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="cellFormat"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        public bool TryGetValueAndFormat(Cell cell, CellFormat cellFormat, out (string value, string formatting) result)
        {
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                // The cell's inner text is an index to shared string table
                int sharedStringIndex = Convert.ToInt32(cell.InnerText, CultureInfo.InvariantCulture);

                // Get the value from the shared string table
                result.value = SharedStringTable[sharedStringIndex];
                result.formatting = string.Empty;

                return true;
            }

            result = default;
            return false;
        }
    }
}
