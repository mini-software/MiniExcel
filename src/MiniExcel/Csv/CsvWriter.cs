using MiniExcelLibs.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace MiniExcelLibs.Csv
{
    internal class CsvWriter : IExcelWriter
    {
	   public void SaveAs(Stream stream, object input)
	   {
		  using (StreamWriter writer = new StreamWriter(stream))
		  {
			 // notice : if first one is null then it can't get Type infomation
			 var first = true;
			 Type type;
			 PropertyInfo[] props = null;
			 foreach (var e in input as IEnumerable)
			 {
				// head
				if (first)
				{
				    first = false;
				    type = e.GetType();
				    props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
				    writer.Write(string.Join(",", props.Select(s => CsvHelpers.ConvertToCsvValue(s.Name))));
				    writer.Write(Environment.NewLine);
				}

				var values = props.Select(s => CsvHelpers.ConvertToCsvValue(s.GetValue(e)?.ToString()));
				writer.Write(string.Join(",", values));
				writer.Write(Environment.NewLine);
			 }
		  }
	   }
    }
}
