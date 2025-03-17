using System;
using System.Globalization;
using System.Linq;

namespace MiniExcelLibs.Utils
{
    internal static class ReferenceHelper
    {
        public static string GetCellNumber(string cell)
        {
            return cell
                .Where(char.IsDigit)
                .Aggregate(string.Empty, (current, c) => current + c);
        }

        public static string GetCellLetter(string cell)
        {
            return cell
                .Where(char.IsLetter)
                .Aggregate(string.Empty, (current, c) => current + c);
        }

        /// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
        public static Tuple<int, int> ConvertCellToXY(string cell)
        {
            const string keys = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            const int mode = 26;

            //AA=27,ZZ=702
            var cellLetter = GetCellLetter(cell);
            var x = cellLetter.Aggregate(0, (idx, chr) => idx * mode + keys.IndexOf(chr));

            var cellNumber = GetCellNumber(cell);
            return Tuple.Create(x, int.Parse(cellNumber));
        }

        /// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
        public static string ConvertXyToCell(int x, int y)
        {
            int dividend = x;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return $"{columnName}{y}";
        }

        /**The code below was copied and modified from ExcelDataReader - @MIT License**/
        /// <summary>
        /// Logic for the Excel dimensions. Ex: A15
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="column">The column, 1-based.</param>
        /// <param name="row">The row, 1-based.</param>
        public static bool ParseReference(string value, out int column, out int row)
        {
            row = 0;
            column = 0;
            var position = 0;
            const int offset = 'A' - 1;

            if (string.IsNullOrWhiteSpace(value))
                return false;

            while (position < value.Length)
            {
                var c = char.ToUpperInvariant(value[position]);
                if (c >= 'A' && c <= 'Z')
                {
                    position++;
                    column *= 26;
                    column += c - offset;
                    continue;
                }

                if (!char.IsLetter(c))
                {
                    if (char.IsDigit(c))
                        break;
                    return false;
                }

                row = 0;
                column = 0;
                position = 0;
                break;
            }

            if (position == 0)
                return false;

            return int.TryParse(value.Substring(position), NumberStyles.None, CultureInfo.InvariantCulture, out row);
        }
    }
}
