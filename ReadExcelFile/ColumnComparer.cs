using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcelFile
{
    class ColumnComparer : IComparer<string>
    {
        private const string DigitsRegex = @"\d+";
        private const string LettersRegex = "[A-Za-z]+";
        private const int AlphabetLength = 'Z' - 'A' + 1;

        public int Compare(string x, string y)
        {
            if (x.Length > y.Length)
            {
                return 1;
            }
            else if (x.Length < y.Length)
            {
                return -1;
            }
            else
            {
                return string.Compare(x, y, true);
            }
        }

        public static Func<Cell, bool> IsColumnInRange(string firstColumn, string lastColumn) => c => IsColumnInRange(firstColumn, lastColumn, GetColumnName(c.CellReference.Value));

        public static bool IsColumnInRange(string firstColumn, string lastColumn, string columnName) => CompareColumn(columnName, firstColumn) >= 0 && CompareColumn(columnName, lastColumn) <= 0;

        // Given a cell name, parses the specified cell to get the row index.
        public static uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(DigitsRegex);
            System.Text.RegularExpressions.Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }
        // Given a cell name, parses the specified cell to get the column name.
        public static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(LettersRegex);
            System.Text.RegularExpressions.Match match = regex.Match(cellName);

            return match.Value;
        }

        // Given two columns, compares the columns.
        public static int CompareColumn(string column1, string column2)
        {
            if (column1.Length > column2.Length)
            {
                return 1;
            }
            else if (column1.Length < column2.Length)
            {
                return -1;
            }
            else
            {
                return string.Compare(column1, column2, true);
            }
        }

        public static string NextColumn(string column)
        {
            char[] cleanColumn = column.ToUpper().ToArray<char>();
            return RecurseIncrementCharAtPos(cleanColumn, cleanColumn.Length - 1);
        }

        public static string RecurseIncrementCharAtPos(char[] column, int pos)
        {
            var letter = column[pos];
            if (letter < 'Z')
            {
                return IncrementCharAtPos(column, pos);
            }
            else
            {
                letter = 'A';
                column[pos] = letter;
                pos--;
                if (pos < 0)
                {
                    var newCleanColumn = new char[column.Length + 1];
                    newCleanColumn[0] = 'A';
                    Array.Copy(column, 0, newCleanColumn, 1, column.Length);
                    return new string(newCleanColumn);
                }
                else
                {
                    return RecurseIncrementCharAtPos(column, pos);
                }
            }
        }

        public static string IncrementCharAtPos(char[] column, int pos)
        {
            char[] cleanColumn = column.ToArray<char>();
            var letter = cleanColumn[pos];
            letter++;
            cleanColumn[pos] = letter;
            return new string(cleanColumn);
        }
        public static string IncrementCharAtPos(string column, int pos)
        {
            char[] cleanColumn = column.ToArray<char>();
            var letter = cleanColumn[pos];
            letter++;
            cleanColumn[pos] = letter;
            return new string(cleanColumn);
        }
    }
}
