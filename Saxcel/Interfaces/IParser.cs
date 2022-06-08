using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace Saxcel
{
    /// <summary>
    /// A parser interface for different datatypes.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface IParser<T>
    {
        /// <summary>
        /// Collection that holds the different formats of data.
        /// </summary>
        Dictionary<int, string> Formats { get; set; }

        /// <summary>
        /// Adds a format to the parsers collection of formats.
        /// </summary>
        /// <param name="customFormat"></param>
        void AddFormat((int key, string value) customFormat);

        /// <summary>
        /// Returns the cell value (of type T) and the formatting string for displaying that value.
        /// </summary>
        /// <param name="fromCell"></param>
        /// <param name="withFormat"></param>
        /// <param name="valueFormat"></param>
        /// <returns></returns>
        bool TryGetValueAndFormat(Cell fromCell, CellFormat withFormat, out (T value, string formatting) valueFormat);
    }
}
