using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace Saxcel
{
    public interface IParser<T>
    {
        Dictionary<int, string> Formats { get; set; }

        void AddFormat((int key, string value) customFormat);

        bool TryGetValueAndFormat(Cell fromCell, CellFormat withFormat, out (T value, string formatting) valueFormat);
    }
}
