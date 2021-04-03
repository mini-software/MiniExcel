using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MiniExcelLibs.Csv
{
    internal static class CsvHelpers
    {
	   /// <summary>If content contains `;, "` then use "{value}" format</summary>
	   public static string ConvertToCsvValue(string value)
	   {
		  if (value == null)
			 return string.Empty;
		  if (value.Contains("\""))
		  {
			 value = value.Replace("\"", "\"\"");
			 return $"\"{value}\"";
		  }
		  else if (value.Contains(",") || value.Contains(" "))
		  {
			 return $"\"{value}\"";
		  }
		  return value;
	   }
    }
}
