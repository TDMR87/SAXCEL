# SAXCEL (Excel file reader using SAX) for .NET C#

A library for reading large spreadsheets (__millions__ of rows) without running out of memory. 

Utilizes the SAX approach (Simple API for XML) but wraps a lot of boilerplate code into a minimalistic and easy to use API.
https://docs.microsoft.com/en-us/office/open-xml/how-to-parse-and-read-a-large-spreadsheet

---

# HOW TO USE

Reference Saxcel
```C#
using Saxcel;
```

Initialize the reader in a using block, specifying the filepath and the sheet to read.

Use a while-loop to iterate through all the cells in a specific column.
```C#
using (var reader = new XlsxReader(filepath: @"C:\largefile.xlsx", sheetname: "Sheet1"))
{
    while (reader.IsReading(column: "B", out string val))
    {
        Console.WriteLine(val);
    }
}
```

Optionally, you can specify a range of columns to read.
```C#
while (reader.IsReading(column: "B", toColumn: "F", out string val))
{
    Console.WriteLine(val);
}
```

You can even specify a range of columns with specific rows.
```C#
while (reader.IsReadingRange("C999:D3420", out string val))
{
    Console.WriteLine(val);
}
```

The XlsxReader exposes a few public properties that tell, for example, on what row and on what column the reader is currently on.
```C#
while (reader.IsReading(column: "A", toColumn: "F", out string val))
{
    if (reader.CurrentColumn == "C" || reader.CurrentRow == 99) continue;
    else Console.WriteLine(val);
}
```
