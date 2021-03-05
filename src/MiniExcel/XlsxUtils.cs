namespace MiniExcelLibs
{
    using System;

    internal static class XlsxUtils
    {
        /// <summary>
        /// Encode to XML (special characteres: &apos; &quot; &gt; &lt; &amp;)
        /// </summary>
        internal static string EncodeXML(object value) => value == null 
                ? "" 
                : value.ToString().Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");

        /// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
        internal static string ConvertXyToCell(Tuple<int, int> xy)
        {
            return ConvertXyToCell(xy.Item1, xy.Item2);
        }

        /// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
        internal static string ConvertXyToCell(int x, int y)
        {
            int dividend = x;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return $"{columnName}{y}";
        }

        /// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
        internal static Tuple<int, int> ConvertCellToXY(string cell)
        {
            const string keys = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            const int mode = 26;

            var x = 0;
            var cellLetter = GetCellLetter(cell);
            //AA=27,ZZ=702
            for (int i = 0; i < cellLetter.Length; i++)
                x = x * mode + keys.IndexOf(cellLetter[i]);

            var cellNumber = GetCellNumber(cell);
            return Tuple.Create(x, int.Parse(cellNumber));
        }

        internal static string GetCellNumber(string cell)
        {
            string cellNumber = string.Empty;
            for (int i = 0; i < cell.Length; i++)
            {
                if (Char.IsDigit(cell[i]))
                    cellNumber += cell[i];
            }
            return cellNumber;
        }

        internal static string GetCellLetter(string cell)
        {
            string GetCellLetter = string.Empty;
            for (int i = 0; i < cell.Length; i++)
            {
                if (Char.IsLetter(cell[i]))
                    GetCellLetter += cell[i];
            }
            return GetCellLetter;
        }
    }
}
