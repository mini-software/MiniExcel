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
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

void Main()
{
	var doc = new XmlDocument();
	doc.LoadXml(xml);
	
	var bytes = Encoding.UTF8.GetBytes(xml);
	using (var stream = new MemoryStream())
	{
		stream.Write(bytes);
		stream.Position=0;

		var reader = XmlReader.Create(stream, _xmlSettings);

		reader.MoveToContent();
		reader.Read();
		while (reader.Read())
		{
			Console.WriteLine($"{reader.Name},{reader.Value}");
		}

//		stream.Position = 0;
//
//		reader = XmlReader.Create(stream, _xmlSettings);
//		reader.MoveToContent();
//		reader.Read();
//		while (reader.Read())
//		{
//			Console.WriteLine($"{reader.Name},{reader.Value}");
//		}
	}



}

private static readonly XmlReaderSettings _xmlSettings = new XmlReaderSettings
{
	IgnoreComments = true,
	IgnoreWhitespace = true,
	XmlResolver = null,
};

// You can define other methods, fields, classes and namespaces here
public class SharingStringReader
{
	// sharingstringreder
}

const string xml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""
    xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x14ac xr xr2 xr3""
    xmlns:x14ac=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac""
    xmlns:xr=""http://schemas.microsoft.com/office/spreadsheetml/2014/revision""
    xmlns:xr2=""http://schemas.microsoft.com/office/spreadsheetml/2015/revision2""
    xmlns:xr3=""http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"" xr:uid=""{7A401454-31EE-4EA0-B820-0475E0C5C4AB}"">
    <dimension ref=""B2:G8""/>
    <sheetViews>
        <sheetView tabSelected=""1"" workbookViewId=""0"">
            <selection activeCell=""G2"" sqref=""G2""/>
        </sheetView>
    </sheetViews>
    <sheetFormatPr defaultRowHeight=""14.4"" x14ac:dyDescent=""0.3""/>
    <sheetData>
        <row r=""2"" spans=""2:7"" x14ac:dyDescent=""0.3"">
            <c r=""C2"">
                <v>1</v>
            </c>
            <c r=""G2"">
                <v>1</v>
            </c>
        </row>
        <row r=""5"" spans=""2:7"" x14ac:dyDescent=""0.3"">
            <c r=""E5"">
                <v>1</v>
            </c>
        </row>
        <row r=""8"" spans=""2:7"" x14ac:dyDescent=""0.3"">
            <c r=""B8"">
                <v>1</v>
            </c>
        </row>
    </sheetData>
    <pageMargins left=""0.7"" right=""0.7"" top=""0.75"" bottom=""0.75"" header=""0.3"" footer=""0.3""/>
    <pageSetup paperSize=""9"" orientation=""portrait"" r:id=""rId1""/>
</worksheet>";