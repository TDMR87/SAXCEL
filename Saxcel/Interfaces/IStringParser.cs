using System.Collections.Generic;

namespace Saxcel
{
    public interface IStringParser : IParser<string>
    {
        Dictionary<int, string> SharedStringTable { get; set; }
    }
}
