using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Saxcel
{
    public class DateParser : IParser<DateTime>
    {
        /// <summary>
        /// Checks that the cell value is date format and returns a ValueTuple<DateTime, string>
        /// that contains the value itself (DateTime type) and the formatting for displaying that
        /// value. 
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
                    result.value = DateTime.FromOADate(dateDouble);
                    result.formatting = Formats[cellFormat.NumberFormatId.AsInt()];
                    return true;
                }
              }

            result = default;
            return false;
        }

        /// <summary>
        /// Adds a DateTime format to the collection of formats.
        /// </summary>
        /// <param name="customFormat"></param>
        public void AddFormat((int key, string value) customFormat)
        {
            // If formatting is for "long date"
            if (customFormat.value.StartsWith("[$-F800]"))
            {
                // Replace parts of the formatting string
                // in order to get the right formatting for long dates in C#
                var trimmedFormatting = customFormat.value
                    .Replace("[$-F800]", "")
                    .Replace("mmmm", "MMMM")
                    .Replace("\\,\\ y", " y")
                    .Replace("\\", "")
                    .Replace("dd ", "d ");

                // Add the trimmed formatting to the collection of formats
                Formats.Add(customFormat.key, trimmedFormatting);
            }
            else
            {
                // Add the formatting to the collection of formats
                Formats.Add(customFormat.key, customFormat.value);
            }
        }

        /// <summary>
        /// Contains date formats and their corresponding id code as the key.
        /// https://msdn.microsoft.com/en-GB/library/documentformat.openxml.spreadsheet.numberingformat(v=office.14).aspx
        /// </summary>
        public Dictionary<int, string> Formats { get; set; } = new Dictionary<int, string>()
        {
            [12] = "# ?/?" ,
            [13] = "# ??/??", 
            [14] = "d/M/yyyy",
            [15] = "d-mmm-yy",
            [16] = "d-mmm",
            [17] = "mmm-yy",
            [18] = "h:mm tt",
            [19] = "h:mm:ss tt",
            [20] = "H:mm",
            [21] = "H:mm:ss",
            [22] = "m/d/yyyy H:mm",
            [37] = "#,##0 ;(#,##0)",
            [38] = "d.M.yy",
            [39] = "yyyy-MM-dd",
            [40] = "dd MMMM yyyy",
            [45] = "d MMMM yyyy",
            [46] = "M/d",
            [47] = "M/d/yy",
            [48] = "MM/dd/yy",
            [49] = "d-MMM",
        };
    }
}
