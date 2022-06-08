using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Saxcel
{
    internal class TimeParser : IParser<DateTime>
    {
        /// <summary>
        /// Checks that the cell value is time format and returns a ValueTuple<DateTime, string>
        /// that contains the value (DateTime type) and formatting. 
        /// </summary>
        /// <param name="cellValue"></param>
        /// <param name="cellFormat"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public bool TryGetValueAndFormat(Cell cell, CellFormat cellFormat, out (DateTime value, string formatting) result)
        {
            if (cellFormat != null && Formats.ContainsKey(cellFormat.NumberFormatId.AsInt()))
            {
                // Parse the value to a double, because in .xlsx files dates are stored as a number of days since 1.1.1900
                if (double.TryParse(cell.CellValue.InnerText, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var dateDouble))
                {
                    // Set and return result
                    result.value = DateTime.FromOADate(dateDouble);
                    result.formatting = Formats[cellFormat.NumberFormatId.AsInt()];
                    return true;
                }
            }

            result = default;
            return false;
        }

        /// <summary>
        /// Contains number formats where the corresponding number format id is the key and the format is the value.
        /// </summary>
        public Dictionary<int, string> Formats { get; set; } = new Dictionary<int, string>()
        {
            /*
             * XlsxReader will add to this collection all the custom Time formattings
             * that are specified in a xlsx workbook.
             */
        };

        /// <summary>
        /// Adds a format to the collection of formats.
        /// </summary>
        /// <param name="customFormat"></param>
        public void AddFormat((int key, string value) customFormat)
        {
            // If formatting is for "time"
            if (customFormat.value.StartsWith("[$-F400]"))
            {
                // Extract the formatting part, e.g. h:mm:ss
                var trimmedFormatting = Regex.Match(customFormat.value, "\\w+:\\w+:\\w+")?.Value;
                Formats.Add(customFormat.key, trimmedFormatting);
            }
            else
            {
                Formats.Add(customFormat.key, customFormat.value);
            }
        }
    }
}
