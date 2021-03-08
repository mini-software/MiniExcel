<Query Kind="Program">
  <NuGetReference>AngleSharp</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>DocumentFormat.OpenXml</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.IO.Compression</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

void Main()
{
	var source = new Dictionary<int, Dictionary<int, object>>()
	{
		{0,new Dictionary<int,object>(){{0,0},{3,3}}},
		{3,new Dictionary<int,object>(){{2,2}}},
	};
	Console.WriteLine(JsonConvert.SerializeObject(source, Newtonsoft.Json.Formatting.Indented));
	var rows = GetRows(@"D:\git\MiniExcel\samples\xlsx\TestCenterEmptyRow\TestCenterEmptyRow.xlsx").ToList();
	Console.WriteLine(rows);
}

private static string ConvertToString(ZipArchiveEntry entry)
{
	if (entry == null)
		return null;
	using (var eStream = entry.Open())
	using (var reader = new StreamReader(eStream))
		return reader.ReadToEnd();
}

internal static class ExcelXName
{
	internal readonly static XNamespace ExcelNamespace = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
	internal readonly static XNamespace ExcelRelationshipsNamepace = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
	internal readonly static XName Row;
	internal readonly static XName R;
	internal readonly static XName V;
	internal readonly static XName T;
	internal readonly static XName C;
	internal readonly static XName Dimension;
	static ExcelXName()
	{
		Row = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "row";
		R = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "r";
		V = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "v";
		T = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "t";
		C = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "c";
		Dimension = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "dimension";
	}
}

IEnumerable<IEnumerable<object>> GetRows(string path)
{
	using (FileStream stream = new FileStream(path, FileMode.Open))
	using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false, UTF8Encoding.UTF8))
	{
		var firstSheetEntry = archive.Entries.First(w => w.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase));
		var xml = ConvertToString(firstSheetEntry);
		var xl = XElement.Parse(xml);

		var maxRowIndex = 3;
		var maxColumnIndex = 3;

		// first get sheet row count & column count

		var rowIndex = 0;
		foreach (var row in xl.Descendants(ExcelXName.Row))
		{
			{
				var r = row.Attribute("r")?.Value?.ToString();

				var rIndex = int.MinValue;
				if (int.TryParse(r, out var _rowIndex))
					rIndex = _rowIndex - 1; // The row attribute is 1 - based			
			}


			var cells = new List<object>(maxColumnIndex);

			foreach (var cell in row.Descendants(ExcelXName.C))
			{

			}

			yield return cells;
		}
	}
}

void Main2()
{
	{
		var dic = new Dictionary<int, Dictionary<int, object>>() {
			{0,new Dictionary<int,object>(){{0,0},{3,3}}},
			{3,new Dictionary<int,object>(){{2,3}}},
		};
		Console.WriteLine(JsonConvert.SerializeObject(dic, Newtonsoft.Json.Formatting.Indented));

		var maxRowIndex = 3;
		var maxColumnIndex = 3;
		for (int rowIndex = 0; rowIndex <= maxRowIndex; rowIndex++)
		{
			if (!dic.ContainsKey(rowIndex))
			{
				var d = new Dictionary<int, object>();
				for (int columnIndex = 0; columnIndex <= maxColumnIndex; columnIndex++)
				{
					d.Add(columnIndex, null);
				}
				dic.Add(rowIndex, d);
			}
			else
			{
				var d = dic[rowIndex];
				for (int columnIndex = 0; columnIndex <= maxColumnIndex; columnIndex++)
				{
					if (!d.ContainsKey(columnIndex))
					{
						d.Add(columnIndex, null);
					}
				}
			}
		}

		Console.WriteLine(JsonConvert.SerializeObject(dic, Newtonsoft.Json.Formatting.Indented));
	}
	{
		var dic = new Dictionary<int, Dictionary<int, object>>() {
			{0,new Dictionary<int,object>(){{0,0},{1,null},{2,null},{3,3}}},
			{1,new Dictionary<int,object>(){{0,null},{1,null},{2,null},{3,null}}},
			{2,new Dictionary<int,object>(){{0,null},{1,null},{2,null},{3,null}}},
			{3,new Dictionary<int,object>(){{0,null},{1,null},{2,2},{3,null}}},
		};
		//Console.WriteLine(dic);
		var json = JsonConvert.SerializeObject(dic, Newtonsoft.Json.Formatting.Indented);
		Console.WriteLine(json);
	}
}

// You can define other methods, fields, classes and namespaces here
