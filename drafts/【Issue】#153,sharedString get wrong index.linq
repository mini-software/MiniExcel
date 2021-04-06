<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Diagnostics</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Text</RemoveNamespace>
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

void Main()
{
	BugTest();
	FixTest();
}

void FixTest()
{
	Console.WriteLine("==== Fix ====");
	using (var stream = File.OpenRead(@"D:\git\MiniExcel\samples\xlsx\TestIssue153.xlsx"))
	using (var ZipArchive = new System.IO.Compression.ZipArchive(stream))
	using (var stream2 = ZipArchive.Entries.Single(w => w.FullName.Contains("shared")).Open())
	{
		var list = GetSharedStrings(stream2).ToList();
		var takes = new[] { 4647, 4648, 4940, 4649, 4655, 0, 1, 4650, 4840, 4653, 8791, 8789, 4654, 2, 4651, 3,4841,4652 };
		foreach (var t in takes)
		{
			Console.WriteLine($"{t} , {list[t]}");
		}
	}
}

internal IEnumerable<string> GetSharedStrings(Stream stream)
{
	using (var reader = XmlReader.Create(stream))
	{
		if (!reader.IsStartElement("sst", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
			yield break;

		if (!XmlReaderHelper.ReadFirstContent(reader))
			yield break;

		while (!reader.EOF)
		{
			if (reader.IsStartElement("si", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
			{
				var value = StringHelper.ReadStringItem(reader);
				yield return value;
			}
			else if (!XmlReaderHelper.SkipContent(reader))
			{
				break;
			}
		}
	}
}

void BugTest()
{
	Console.WriteLine("==== Bug ====");
	using (var stream = File.OpenRead(@"D:\git\MiniExcel\samples\xlsx\TestIssue153.xlsx"))
	using (var ZipArchive = new System.IO.Compression.ZipArchive(stream))
	using (var stream2 = ZipArchive.Entries.Single(w=>w.FullName.Contains("shared")).Open())
	{
		var xl = XElement.Load(stream2);
		var ts = xl.Descendants(ExcelOpenXmlXName.T).Select((s, i) => new { i, v = s.Value?.ToString() })
			  .ToDictionary(s => s.i, s => s.v)
		;//TODO:need recode

		var takes = new[] { 4647, 4648, 4940, 4649, 4655, 0, 1, 4650, 4840, 4653, 8791, 8789, 4654, 2, 4651, 3,4841,4652 };
		foreach (var t in takes)
		{
			Console.WriteLine($"{t} , {ts[t]}");
		}
	}
}

internal static class XmlReaderHelper
{
	public static bool ReadFirstContent(XmlReader xmlReader)
	{
		if (xmlReader.IsEmptyElement)
		{
			xmlReader.Read();
			return false;
		}

		xmlReader.MoveToContent();
		xmlReader.Read();
		return true;
	}

	public static bool SkipContent(XmlReader xmlReader)
	{
		if (xmlReader.NodeType == XmlNodeType.EndElement)
		{
			xmlReader.Read();
			return false;
		}

		xmlReader.Skip();
		return true;
	}
}

// You can define other methods, fields, classes and namespaces here
internal static class ExcelOpenXmlXName
{
	internal readonly static XNamespace ExcelNamespace = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
	internal readonly static XNamespace ExcelRelationshipsNamepace = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
	internal readonly static XName Row;
	internal readonly static XName R;
	internal readonly static XName V;
	internal readonly static XName T;
	internal readonly static XName C;
	internal readonly static XName Dimension;
	internal readonly static XName Sheet;
	static ExcelOpenXmlXName()
	{
		Row = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "row";
		R = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "r";
		V = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "v";
		T = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "t";
		C = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "c";
		Dimension = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "dimension";
		Sheet = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "sheet";
	}
}

internal static class StringHelper
{
	private const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

	private const string ElementT = "t";
	private const string ElementR = "r";

	public static string ReadStringItem(XmlReader reader)
	{
		string result = string.Empty;
		if (!XmlReaderHelper.ReadFirstContent(reader))
		{
			return result;
		}

		while (!reader.EOF)
		{
			if (reader.IsStartElement(ElementT, NsSpreadsheetMl))
			{
				// There are multiple <t> in a <si>. Concatenate <t> within an <si>.
				result += reader.ReadElementContentAsString();
			}
			else if (reader.IsStartElement(ElementR, NsSpreadsheetMl))
			{
				result += ReadRichTextRun(reader);
			}
			else if (!XmlReaderHelper.SkipContent(reader))
			{
				break;
			}
		}

		return result;
	}

	private static string ReadRichTextRun(XmlReader reader)
	{
		string result = string.Empty;
		if (!XmlReaderHelper.ReadFirstContent(reader))
		{
			return result;
		}

		while (!reader.EOF)
		{
			if (reader.IsStartElement(ElementT, NsSpreadsheetMl))
			{
				result += reader.ReadElementContentAsString();
			}
			else if (!XmlReaderHelper.SkipContent(reader))
			{
				break;
			}
		}

		return result;
	}
}