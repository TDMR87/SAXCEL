using System.Collections.Generic;

namespace Saxcel
{
    /// <summary>
    /// A parser interface for string types.
    /// </summary>
    public interface IStringParser : IParser<string>
    {
        /// <summary>
        /// The shared string table is a key-value dictionary containing
        /// the text values of an .xlsx file.
        /// </summary>
        Dictionary<int, string> SharedStringTable { get; set; }
    }
}
