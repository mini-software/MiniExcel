<Query Kind="Program">
  <NuGetReference>CsvHelper</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Xunit</Namespace>
  <Namespace>CsvHelper</Namespace>
  <Namespace>System.Globalization</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Diagnostics</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Text</RemoveNamespace>
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

#load "xunit"

[Fact()]
void Main()
{
	{
		{
			var input = Enumerable.Range(1, 10).Select((s, idx) => new { Id = idx, Text = Guid.NewGuid().ToString() }).ToList();
			var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.csv");
			CsvImpl.SaveAs(path, "");
			Console.WriteLine(path);
		}
	}

	//TODO:Datetime have default format like yyyy-MM-dd HH:mm:ss ?
	{
		Assert.Equal(Generate("\"\"\""),MiniExcelGenerateCsv("\"\"\""));
		Assert.Equal(Generate(","),MiniExcelGenerateCsv(","));
		Assert.Equal(Generate(" "),MiniExcelGenerateCsv(" "));
		Assert.Equal(Generate(";"),MiniExcelGenerateCsv(";"));
		Assert.Equal(Generate("\t"),MiniExcelGenerateCsv("\t"));
	}
}

string Generate(string value)
{
	var records = Enumerable.Range(1, 1).Select((s, idx) => new { v1 = value, v2 = value });
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.csv");
	using (var writer = new StreamWriter(path))
	using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
	{
		csv.WriteRecords(records);
	}
	Console.WriteLine(path);
	
	return File.ReadAllText(path);
}

string MiniExcelGenerateCsv(string value)
{
	var records = Enumerable.Range(1, 1).Select((s, idx) => new { v1 = value, v2 = value });
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.csv");
	
	using (var stream = File.Create(path))
	{
		stream.SaveAs(records);
	}
	Console.WriteLine(path);
	
	return File.ReadAllText(path);
}

internal static partial class CsvImpl
{
	internal static void SaveAs(this Stream stream, object input)
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
	internal static void SaveAs(string path, object input)
	{
		using (var stream = File.Create(path))
		{
			stream.SaveAs(input);
		}
	}
}

internal static class CsvHelpers
{
	/// <summary>If content contains `;, "` then use "{value}" format</summary>
	public static string ConvertToCsvValue(string value)
	{
		if (value == null)
			return "";
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


// You can define other methods, fields, classes and namespaces here

