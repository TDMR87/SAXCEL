using System;
using Saxcel;

namespace SaxcelConsoleClient1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                using (XlsxReader reader = new XlsxReader(@"C:\temp\file_example_5000_rows.xlsx", "Sheet1"))
                {
                    while (reader.IsReading("A", out string value))
                    {
                        Console.WriteLine(value);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            Console.ReadKey();
        }
    }
}
