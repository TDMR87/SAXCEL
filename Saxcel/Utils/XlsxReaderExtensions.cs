using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Saxcel
{
    internal static class XlsxReaderExtensions
    {
        /// <summary>
        /// Returns the alphabetical letter(s) of the column the cell is in.
        /// For example, if the specified cell is in column "A1", this method returns "A".
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string GetColumnName(this Cell cell)
        {
            string output = "";

            char[] chars = cell.CellReference.Value?.ToCharArray();

            foreach (char c in chars)
            {
                if (char.IsNumber(c))
                {
                    break;
                }

                output += c;
            }

            return output.ToUpper();
        }

        /// <summary>
        /// When an excel column name is specified, e.g. AB, 
        /// this method returns the next column name, e.g. AC
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        internal static string GetNextColumn(this string input)
        {
            // Sanity
            input = input.ToUpper();

            // Builder for the next column letters
            StringBuilder nextColumn = new StringBuilder(input);

            // If just one character in the input
            if (input.Length == 1)
            {
                // Replace the character with the next character of the alphabet
                nextColumn.Replace(input[0], GetNextCharacter(input[0]), 0, 1);

                // Was it the last letter?
                if (input[0] == 'Z')
                {
                    // Also add letter 'A' at the beginning, because after column Z, next column is AA)
                    nextColumn.Insert(0, 'A');
                }
            }
            else // If more characters, e.g. column AAZ -> return ABA
            {
                // Iterate the input string from end to start
                for (int i = input.Length - 1; i >= 0; i--)
                {
                    char currentCharacter = input[i];

                    nextColumn.Replace(currentCharacter, GetNextCharacter(currentCharacter), i, 1);

                    // If the current character is the last alphabet letter
                    if (currentCharacter == 'Z')
                    {
                        // Increment the left side character
                        char leftChar = input[i - 1];
                        nextColumn.Replace(leftChar, GetNextCharacter(leftChar), i - 1, 1);
                    }
                    else
                    {
                        break;
                    }
                }
            }

            // Return the next column
            return nextColumn.ToString().ToUpper();
        }

        /// <summary>
        /// Returns the next character in the alphabet.
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        internal static char GetNextCharacter(this char input)
        {
            return (input == 'Z' ? 'A' : (char)(input + 1));
        }

        /// <summary>
        /// Gets the cell formatting object or null if the format could not be determined.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        internal static bool TryGetCellFormat(this WorkbookPart workbookPart, Cell cell, out CellFormat cellFormat)
        {
            if (cell.StyleIndex != null && int.TryParse(cell.StyleIndex.InnerText, out int styleIndex))
            {
                cellFormat = workbookPart
                    .WorkbookStylesPart
                    .Stylesheet
                    .CellFormats
                    .ChildElements[styleIndex]
                    as CellFormat;

                return true;
            }

            cellFormat = null;
            return false;
        }

        /// <summary>
        /// Gets the custom formats defined in the workbook's styles.xml.
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <returns></returns>
        internal static IEnumerable<(int formatId, string formatCode)> GetCustomCellFormattings(this WorkbookPart workbookPart)
        {
            var stylePart = workbookPart.WorkbookStylesPart;

            var numFormatsParentNodes = stylePart.Stylesheet.ChildElements.OfType<NumberingFormats>();

            foreach (var numFormatParentNode in numFormatsParentNodes)
            {
                var formatNodes = numFormatParentNode.ChildElements.OfType<NumberingFormat>();

                foreach (var formatNode in formatNodes)
                {
                    yield return (formatNode.NumberFormatId.AsInt(), formatNode.FormatCode.InnerText);
                }
            }
        }

        internal static int AsNumber(this string text)
        {
            string output = string.Empty;

            foreach (char c in text)
            {
                output += ((int)c);
            }

            return int.Parse(output);
        }

        internal static (string, string) GetCellRange(this string range)
        {
            if (Regex.IsMatch(range, @"^[a-zA-Z]+[0-9]{0,}:[a-zA-Z]+[0-9]{0,}$"))
            {
                var split = range.Split(':');
                return (split[0], split[1]);
            }
            else
            {
                throw new ArgumentException("The range was not in a valid format. e.g. A1:B100.", nameof(range));
            }
        }

        /// <summary>
        /// Determines if the cell is in range, e.g. if range is B11:C20 and cell is A6, then the cell is not in range.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        internal static bool IsInRange(this Cell cell, string range)
        {
            var cellRef = SplitAlphabetsAndNumbers(cell.CellReference);
            var cellColumnRef = cellRef.First();
            var cellRowRef = cellRef.Last();

            var rangeFrom = range.Split(':')[0];
            var rangeFromColumn = SplitAlphabetsAndNumbers(rangeFrom).First();
            var rangeFromRow = SplitAlphabetsAndNumbers(rangeFrom).Last();

            var rangeTo = range.Split(':')[1];
            var rangeToColumn = SplitAlphabetsAndNumbers(rangeTo).First();
            var rangeToRow = SplitAlphabetsAndNumbers(rangeTo).Last();

            // If cell is in correct column range
            if (cellColumnRef.AsNumber() >= rangeFromColumn.AsNumber() && cellColumnRef.AsNumber() <= rangeToColumn.AsNumber())
            {
                // If cell is in the first column of the range and row is too low
                if (cellColumnRef == rangeFromColumn && int.Parse(cellRowRef) < int.Parse(rangeFromRow))
                {
                    // not in range
                    return false;
                }

                // If cell is in the last column of the range and row is too high
                if (cellColumnRef == rangeToColumn && int.Parse(cellRowRef) > int.Parse(rangeToRow))
                {
                    // not in range
                    return false;
                }

                // In range
                return true;
            }

            // not in range
            return false;
        }

        public static IEnumerable<string> SplitAlphabetsAndNumbers(this string input)
        {
            var words = new List<string> { string.Empty };
            for (var i = 0; i < input.Length; i++)
            {
                words[words.Count - 1] += input[i];
                if (i + 1 < input.Length && char.IsLetter(input[i]) != char.IsLetter(input[i + 1]))
                {
                    words.Add(string.Empty);
                }
            }
            return words;
        }

        public static int AsInt<T>(this T input)
        {
            try
            {
                return Convert.ToInt32(input.ToString());
            }
            catch (Exception)
            {
                throw new ArgumentException("Converting input to integer failed.", nameof(input));
            }
        }

        public static bool IsOneOf(this string source, params string[] list)
        {
            if (null == source) throw new ArgumentNullException(nameof(source));
            return list.Contains(source, StringComparer.OrdinalIgnoreCase);
        }
    }
}
