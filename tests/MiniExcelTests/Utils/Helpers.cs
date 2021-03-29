/**
 This Class Modified from ExcelDataReader : https://github.com/ExcelDataReader/ExcelDataReader
 **/
namespace MiniExcelLibs.Tests.Utils
{
    using System;
    using System.Collections.Generic;
    using System.Dynamic;
    using System.Globalization;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Reflection;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Xml.Linq;

    internal static class Helpers
    {
	   private static Dictionary<int, string> _IntMappingAlphabet = new Dictionary<int, string>();
	   private static Dictionary<string, int> _AlphabetMappingInt = new Dictionary<string, int>();
	   static Helpers()
	   {
		  for (int i = 0; i <= 255; i++)
		  {
			 _IntMappingAlphabet.Add(i, IntToLetters(i));
			 _AlphabetMappingInt.Add(IntToLetters(i), i);
		  }
	   }

	   public static string GetAlphabetColumnName(int ColumnIndex) => _IntMappingAlphabet[ColumnIndex];
	   public static int GetColumnIndex(string columnName) => _AlphabetMappingInt[columnName];

	   internal static string IntToLetters(int value)
	   {
		  value = value + 1;
		  string result = string.Empty;
		  while (--value >= 0)
		  {
			 result = (char)('A' + value % 26) + result;
			 value /= 26;
		  }
		  return result;
	   }


	   internal static string GetFirstSheetDimensionRefValue(string path)
	   {
		  string refV;
		  using (var stream = File.OpenRead(path))
		  using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false, Encoding.UTF8))
		  {
			 var sheet = archive.Entries.Single(w => w.FullName.StartsWith("xl/worksheets/sheet1", StringComparison.OrdinalIgnoreCase));
			 using (var sheetStream = sheet.Open())
			 {
				var dimension = XElement.Load(sheetStream)
					.Descendants("dimension");
				refV = dimension.Single().Attribute("ref").Value;
			 }
		  }

		  return refV;
	   }
    }

}
