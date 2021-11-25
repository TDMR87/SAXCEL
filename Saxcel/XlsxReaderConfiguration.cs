using System;

namespace Saxcel
{
    public class XlsxReaderConfiguration
    {
        public IStringParser StringParser { get; set; }
        public IParser<DateTime> DateParser { get; set; }
        public IParser<DateTime> TimeParser { get; set; }
        public IParser<decimal> NumberParser { get; set; }
    }
}
