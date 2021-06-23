namespace MiniExcelLibs.Utils
{
    using System;
    using System.Globalization;

    internal static partial class ReferenceHelper
    {
	   public static string GetCellNumber(string cell)
	   {
		  string cellNumber = string.Empty;
		  for (int i = 0; i < cell.Length; i++)
		  {
			 if (Char.IsDigit(cell[i]))
				cellNumber += cell[i];
		  }
		  return cellNumber;
	   }

	   public static string GetCellLetter(string cell)
	   {
		  string GetCellLetter = string.Empty;
		  for (int i = 0; i < cell.Length; i++)
		  {
			 if (Char.IsLetter(cell[i]))
				GetCellLetter += cell[i];
		  }
		  return GetCellLetter;
	   }

	   /// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
	   public static Tuple<int, int> ConvertCellToXY(string cell)
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

	   /// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
	   public static string ConvertXyToCell(int x, int y)
	   {
		  int dividend = x;
		  string columnName = String.Empty;
		  int modulo;

		  while (dividend > 0)
		  {
			 modulo = (dividend - 1) % 26;
			 columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
			 dividend = (int)((x - modulo) / 26);
		  }
		  return $"{columnName}{y}";
	   }
    }

    
    internal static partial class ReferenceHelper
    {
	   /**Below Code Copy&Modified from ExcelDataReader @MIT License**/
	   /// <summary>
	   /// Logic for the Excel dimensions. Ex: A15
	   /// </summary>
	   /// <param name="value">The value.</param>
	   /// <param name="column">The column, 1-based.</param>
	   /// <param name="row">The row, 1-based.</param>
	   public static bool ParseReference(string value, out int column, out int row)
        {
            column = 0;
            var position = 0;
            const int offset = 'A' - 1;

            if (value != null)
            {
                while (position < value.Length)
                {
                    var c = value[position];
                    if (c >= 'A' && c <= 'Z')
                    {
                        position++;
                        column *= 26;
                        column += c - offset;
                        continue;
                    }

                    if (char.IsDigit(c))
                        break;

                    position = 0;
                    break;
                }
            }

            if (position == 0)
            {
                column = 0;
                row = 0;
                return false;
            }

            if (!int.TryParse(value.Substring(position), NumberStyles.None, CultureInfo.InvariantCulture, out row))
            {
                return false;
            }

            return true;
        }
    }

}
