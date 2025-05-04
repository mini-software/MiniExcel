using System;
using System.Runtime.CompilerServices;
using MiniExcelLibs.Utils;

namespace MiniExcelLibs.OpenXml
{
    internal static class ExcelOpenXmlUtils
    {
        public static string MinifyXml(string xml) => xml
            .Replace("\r", "")
            .Replace("\n", "")
            .Replace("\t", "")
            .Trim();

        /// <summary>
        /// Encode to XML (special characteres: &apos; &quot; &gt; &lt; &amp;)
        /// </summary>
        public static string EncodeXML(string value) => value == null ? string.Empty 
            : XmlEncoder.EncodeString(value)
                  .Replace("&", "&amp;")
                  .Replace("<", "&lt;")
                  .Replace(">", "&gt;")
                  .Replace("\"", "&quot;")
                  .Replace("'", "&apos;")
                  .ToString();
        
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

        /// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
        public static string ConvertXyToCell(Tuple<int, int> xy) => ConvertXyToCell(xy.Item1, xy.Item2);

        /// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
        public static Tuple<int, int> ConvertCellToXY(string cell) => Tuple.Create(GetCellColumnIndex(cell), GetCellRowNumber(cell));

        public static int GetColumnNumber(string name)
        {
            int number = -1;
            int pow = 1;
            for (int i = name.Length - 1; i >= 0; i--)
            {
                number += (name[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }

        public static int GetCellColumnIndex(string cell)
        {
            const string keys = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            const int mode = 26;

            var x = 0;
            var cellLetter = GetCellColumnLetter(cell);
            //AA=27,ZZ=702
            foreach (var t in cellLetter)
                x = x * mode + keys.IndexOf(t);

            return x;
        }

        public static int GetCellRowNumber(string cell)
        {
            if (string.IsNullOrEmpty(cell))
                throw new Exception("cell is null or empty");
            var cellNumber = string.Empty;
            foreach (var t in cell)
            {
                if (char.IsDigit(t))
                    cellNumber += t;
            }
            return int.Parse(cellNumber);
        }

        public static string GetCellColumnLetter(string cell)
        {
            string GetCellLetter = string.Empty;
            foreach (var t in cell)
            {
                if (char.IsLetter(t))
                    GetCellLetter += t;
            }
            return GetCellLetter;
        }

        public static string ConvertColumnName(int x)
        {
            int dividend = x;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }
    }
}
