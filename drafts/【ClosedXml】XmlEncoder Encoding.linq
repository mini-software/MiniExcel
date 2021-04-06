<Query Kind="Program">
  <NuGetReference>ClosedXML</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.IO.Compression</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <Namespace>ClosedXML.Excel</Namespace>
</Query>

void Main()
{
	Create();
	
	Console.WriteLine(XmlEncoder.EncodeString("\u0001 \u0002 \u0003 \u0004"));
	Console.WriteLine(XmlEncoder.DecodeString("_x0001_ _x0002_ _x0003_ _x0004_"));

	Test1();
	Test2();
}

void Create()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
	Console.WriteLine(path);
	using (var workbook = new XLWorkbook())
	{
		char[] chars = new char[] {'\u0000','\u0001','\u0002','\u0003','\u0004','\u0005','\u0006','\u0007','\u0008',
				'\u0009', //<HT>
	            '\u000A', //<LF>
	            '\u000B','\u000C',
				 '\u000D', //<CR>
	            '\u000E','\u000F','\u0010','\u0011','\u0012','\u0013','\u0014','\u0015','\u0016',
				 '\u0017','\u0018','\u0019','\u001A','\u001B','\u001C','\u001D','\u001E','\u001F','\u007F'
			};
		var input = chars.Select(s => new { Test = s.ToString() });

		var worksheet = workbook.Worksheets.Add("Sample Sheet");

		var index = 1;
		worksheet.Cell($"A{index++}").Value = null;
		foreach (var c in chars)
		{
			worksheet.Cell($"A{index++}").Value = c;
		}
		
		workbook.SaveAs(path);
	}
}

void Test1()
{
	var input = Enumerable.Range(1, 10).Select(s => new { Test1 = "\u0006", Test2 = "123" });
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

	MiniExcelLibs.MiniExcel.SaveAs(path, input);

	Console.WriteLine(path);
}

void Test2()
{
	var input = Enumerable.Range(1, 10).Select(s => new { Test1 = XmlEncoder.EncodeString("\u0006"), Test2 = "123" });
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

	MiniExcelLibs.MiniExcel.SaveAs(path, input);

	Console.WriteLine(path);
}

/// <summary>Class from https://github.com/ClosedXML</summary>
public static class XmlEncoder
{
	private static readonly Regex xHHHHRegex = new Regex("_(x[\\dA-Fa-f]{4})_", RegexOptions.Compiled);
	private static readonly Regex Uppercase_X_HHHHRegex = new Regex("_(X[\\dA-Fa-f]{4})_", RegexOptions.Compiled);

	public static string EncodeString(string encodeStr)
	{
		if (encodeStr == null) return null;

		encodeStr = xHHHHRegex.Replace(encodeStr, "_x005F_$1_");

		var sb = new StringBuilder(encodeStr.Length);

		foreach (var ch in encodeStr)
		{
			if (XmlConvert.IsXmlChar(ch))
				sb.Append(ch);
			else
				sb.Append(XmlConvert.EncodeName(ch.ToString()));
		}

		return sb.ToString();
	}

	public static string DecodeString(string decodeStr)
	{
		if (string.IsNullOrEmpty(decodeStr))
			return string.Empty;
		decodeStr = Uppercase_X_HHHHRegex.Replace(decodeStr, "_x005F_$1_");
		return XmlConvert.DecodeName(decodeStr);
	}
}