/**
 This Class Modified from ExcelDataReader : https://github.com/ExcelDataReader/ExcelDataReader
 **/
namespace MiniExcelLibs.Utils
{
    using System;
    using System.Collections.Generic;
    using System.Dynamic;
    using System.Globalization;
    using System.Linq;
    using System.Reflection;
    using System.Text.RegularExpressions;
    
    internal static class Helpers
    {
	   private static readonly Regex EscapeRegex = new Regex("_x([0-9A-F]{4,4})_");
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
	   public static IDictionary<string, object> GetEmptyExpandoObject(int maxColumnIndex)
	   {
		  // TODO: strong type mapping can ignore this
		  // TODO: it can recode better performance 
		  var cell = (IDictionary<string, object>)new ExpandoObject();
		  for (int i = 0; i <= maxColumnIndex; i++)
            {
			 var key = GetAlphabetColumnName(i);
			 if (!cell.ContainsKey(key))
				cell.Add(key, null);
		  }
		  return cell;
	   }

	   public static IDictionary<string, object> GetEmptyExpandoObject(Dictionary<int, string> hearrows)
	   {
		  // TODO: strong type mapping can ignore this
		  // TODO: it can recode better performance 
		  var cell = (IDictionary<string, object>)new ExpandoObject();
		  foreach (var hr in hearrows)
			 if(!cell.ContainsKey(hr.Value))
				cell.Add(hr.Value, null);
		  return cell;
	   }

	   public static IEnumerable<PropertyInfo> GetPropertiesWithSetter(Type type)
	   {
		  return type.GetProperties(BindingFlags.SetProperty |
					BindingFlags.Public |
					BindingFlags.Instance).Where(prop => prop.GetSetMethod() != null);
	   }


	   public static string ConvertEscapeChars(string input)
	   {
		  return EscapeRegex.Replace(input, m => ((char)uint.Parse(m.Groups[1].Value, NumberStyles.HexNumber)).ToString());
	   }

	   /// <summary>
	   /// Convert a double from Excel to an OA DateTime double. 
	   /// The returned value is normalized to the '1900' date mode and adjusted for the 1900 leap year bug.
	   /// </summary>
	   public static double AdjustOADateTime(double value, bool date1904)
	   {
		  if (!date1904)
		  {
			 // Workaround for 1900 leap year bug in Excel
			 if (value >= 0.0 && value < 60.0)
				return value + 1;
		  }
		  else
		  {
			 return value + 1462.0;
		  }

		  return value;
	   }

	   public static bool IsValidOADateTime(double value)
	   {
		  return value > DateTimeHelper.OADateMinAsDouble && value < DateTimeHelper.OADateMaxAsDouble;
	   }

	   public static object ConvertFromOATime(double value, bool date1904)
	   {
		  var dateValue = AdjustOADateTime(value, date1904);
		  if (IsValidOADateTime(dateValue))
			 return DateTimeHelper.FromOADate(dateValue);
		  return value;
	   }
    }

}
