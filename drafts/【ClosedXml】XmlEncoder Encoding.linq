<Query Kind="Program">
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.IO.Compression</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

void Main()
{
	Test1();
	Test2();
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