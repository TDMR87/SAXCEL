using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Saxcel
{
    public class NumberParser : IParser<decimal>
    {
        public string Formatting { get; set; }

        public bool TryGetValueAndFormat(Cell cell, CellFormat cellFormat, out (decimal value, string formatting) result)
        {
            if ((cell.DataType != null && cell.DataType == CellValues.Number) || (cellFormat != null && Formats.ContainsKey(cellFormat.NumberFormatId.AsInt())))
            {
                // Set and return result
                result.value = decimal.Parse(
                               cell.CellValue.InnerText, 
                               NumberStyles.AllowExponent | NumberStyles.AllowDecimalPoint,
                               CultureInfo.InvariantCulture);

                if (cellFormat != null && Formats.ContainsKey(cellFormat.NumberFormatId.AsInt()))
                    result.formatting = Formats[cellFormat.NumberFormatId.AsInt()];
                else
                    result.formatting = "";

                return true;
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
             * Initialize with default number formats.
             * https://msdn.microsoft.com/en-GB/library/documentformat.openxml.spreadsheet.numberingformat(v=office.14).aspx
             * 
             * XlsxReader will add to this collection all the custom formattings
             * that are specified in a xlsx workbook.
             */

            [1] = "0",
            [2] = "0.00",
            [3] = "#,##0",
            [9] = "0%",
            [10] = "0.00%",
            [11] = "0.00E+00"
        };

        public void AddFormat((int key, string value) customFormat)
        {
            if (customFormat.value.Contains(";"))
            {
                // Trim and remove excess characters from the formatting code
                string trimmedValue = customFormat.value.Split(';')[0];
                trimmedValue = trimmedValue.Trim(new char[] { '_', '-', '*', ' ' });

                Formats.Add(customFormat.key, trimmedValue);
            }
            else
            {
                Formats.Add(customFormat.key, customFormat.value);
            }
        }
    }
}
