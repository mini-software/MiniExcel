namespace MiniExcelLibs.OpenXml
{
    using System;
    internal static class ExcelOpenXmlUtils
    {
	   /// <summary>
	   /// Encode to XML (special characteres: &apos; &quot; &gt; &lt; &amp;)
	   /// </summary>
	   internal static string EncodeXML(object value) => value == null
			   ? string.Empty
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
		  return Tuple.Create(GetCellColumnIndex(cell), GetCellRowNumber(cell));
	   }

	   internal static int GetColumnNumber(string name)
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

	   internal static int GetCellColumnIndex(string cell)
	   {
		  const string keys = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		  const int mode = 26;

		  var x = 0;
		  var cellLetter = GetCellColumnLetter(cell);
		  //AA=27,ZZ=702
		  for (int i = 0; i < cellLetter.Length; i++)
			 x = x * mode + keys.IndexOf(cellLetter[i]);

		  return x;
	   }

	   internal static int GetCellRowNumber(string cell)
	   {
		  if (string.IsNullOrEmpty(cell))
			 throw new Exception("cell is null or empty");
		  string cellNumber = string.Empty;
		  for (int i = 0; i < cell.Length; i++)
		  {
			 if (Char.IsDigit(cell[i]))
				cellNumber += cell[i];
		  }
		  return int.Parse(cellNumber);
	   }

	   internal static string GetCellColumnLetter(string cell)
	   {
		  string GetCellLetter = string.Empty;
		  for (int i = 0; i < cell.Length; i++)
		  {
			 if (Char.IsLetter(cell[i]))
				GetCellLetter += cell[i];
		  }
		  return GetCellLetter;
	   }

	   internal static string ConvertColumnName(int x)
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
		  return columnName;
	   }
    }
}
