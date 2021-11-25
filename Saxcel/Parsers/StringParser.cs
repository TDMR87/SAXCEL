using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Saxcel
{
    internal class StringParser : IStringParser
    {
        public Dictionary<int, string> Formats { get; set; } = new Dictionary<int, string>();

        public Dictionary<int, string> SharedStringTable { get; set; } = new Dictionary<int, string>();

        public void AddFormat((int key, string value) customFormat)
        {
            Formats.Add(customFormat.key, customFormat.value);
        }

        public bool TryGetValueAndFormat(Cell cell, CellFormat cellFormat, out (string value, string formatting) result)
        {
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                // The cell's inner text is an index to shared string table
                int sharedStringIndex = Convert.ToInt32(cell.InnerText, CultureInfo.InvariantCulture);

                // Get the value in the specified index from the shared string table
                result.value = SharedStringTable[sharedStringIndex];
                result.formatting = "";

                return true;
            }

            result = default;
            return false;
        }
    }
}
