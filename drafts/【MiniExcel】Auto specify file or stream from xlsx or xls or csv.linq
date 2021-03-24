<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
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

void Main()
{
	var files = Directory.GetFiles(@"D:\git\ExcelDataReader\test\Resources");
	foreach (var path in files)
	{
		if (Helpers.GetExcelType(path) != Helpers.GetExcelType(File.OpenRead(path)))
			Console.WriteLine($"{path} : {Helpers.GetExcelType(path)} and {Helpers.GetExcelType(File.OpenRead(path))}");
	}
}

// You can define other methods, fields, classes and namespaces here
internal static class Helpers
{
	public enum ExcelType
	{
		XLSX,
		XLS,
		CSV,
		UNKNOWN
	}

	internal static ExcelType GetExcelType(string path)
	{
		switch (Path.GetExtension(path).ToLowerInvariant())
		{
			case ".csv":
				return ExcelType.CSV;
			case ".xlsx":
				return ExcelType.XLSX;
			case ".xls":
				return ExcelType.XLS;
			default:
				return ExcelType.UNKNOWN;
		}
	}
	
	// modified from : [.net - How to know stream is xlsx or xls or csv? - Stack Overflow](https://stackoverflow.com/questions/66731497/how-to-know-stream-is-xlsx-or-xls-or-csv/66765911#66765911)
	internal static ExcelType GetExcelType(Stream stream)
	{
		var buffer = new byte[512];
		stream.Read(buffer, 0, buffer.Length);
		var flag = BitConverter.ToUInt32(buffer, 0);
		switch (flag)
		{
			// Old office format (can be any office file)
			case 0xE011CFD0: 
				return ExcelType.XLS;
			// New office format (can be any ZIP archive)
			case 0x04034B50: 
				return ExcelType.XLSX;
			default :
				return ExcelType.CSV;
		}
	}
}