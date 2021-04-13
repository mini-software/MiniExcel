<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

void Main()
{
	Test();
}

void Test(XmlWriterSettings xmlWriterSettings=null)
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xml");
	Console.WriteLine(path);
	var ws = xmlWriterSettings ?? new XmlWriterSettings {Indent = true, Encoding = Encoding.UTF8, OmitXmlDeclaration = true};
	
	var doc = new XmlDocument();
	doc.LoadXml(xml);
	
	using (var reader = XmlReader.Create(new StringReader(xml)))
	using (var stream = File.Create(path))
	//using (var writer = new StreamWriter(stream, Encoding.UTF8))
	using (var writer = XmlWriter.Create(stream))
	{
		while (reader.Read())
		{
			Console.WriteLine("rdr.NodeType = " + reader.NodeType);
			string elementName = reader.Name; 
			var t = reader.NodeType;
			var e = reader.ReadOuterXml();
			var e2 = reader.ReadInnerXml();

			if (!reader.IsStartElement("worksheet", _ns))
				break;
			if (!XmlReaderHelper.ReadFirstContent(reader))
			 	break;
				
			writer.WriteNode(reader, true);	
			while (!reader.EOF)
			{
				if (reader.IsStartElement("sheetData", _ns))
				{
					writer.WriteNode(reader, true);	
					if (!XmlReaderHelper.ReadFirstContent(reader))
						continue;
					while (!reader.EOF)
					{
						//TODO: if contain {{}} format then 
						if (reader.IsStartElement("row", _ns))
						{
							if (!XmlReaderHelper.ReadFirstContent(reader))
								continue;

							//Cells
							{
								var cellIndex = -1;
								while (!reader.EOF)
								{
									if (reader.IsStartElement("c", _ns))
									{
										cellIndex++;
									}

									if (!XmlReaderHelper.SkipContent(reader))
										break;
								}
							}
						}
						else if (!XmlReaderHelper.SkipContent(reader))
						{
							break;
						}
					}
				}
				else if (!XmlReaderHelper.SkipContent(reader))
				{
					break;
				}
			}

			//var e3 = reader.
			
		}
		
	}
}
private const string _ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
/*
learn from
[Read and write with XmlReader and XmlWriter : Xml Write « XML « C# / CSharp Tutorial](http://www.java2s.com/Tutorial/CSharp/0540__XML/ReadandwritewithXmlReaderandXmlWriter.htm)
*/

// You can define other methods, fields, classes and namespaces here
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

const string xml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""
    xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x14ac xr xr2 xr3""
    xmlns:x14ac=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac""
    xmlns:xr=""http://schemas.microsoft.com/office/spreadsheetml/2014/revision""
    xmlns:xr2=""http://schemas.microsoft.com/office/spreadsheetml/2015/revision2""
    xmlns:xr3=""http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"" xr:uid=""{00000000-0001-0000-0000-000000000000}"">
    <dimension ref=""A1:B6"" />
    <sheetViews>
        <sheetView tabSelected=""1"" zoomScaleNormal=""100"" workbookViewId=""0"">
            <selection activeCell=""C8"" sqref=""C8"" />
        </sheetView>
    </sheetViews>
    <sheetFormatPr defaultColWidth=""11.5546875"" defaultRowHeight=""13.2"" x14ac:dyDescent=""0.25"" />
    <cols>
        <col min=""1"" max=""1"" width=""17.109375"" customWidth=""1"" />
        <col min=""2"" max=""2"" width=""21.77734375"" customWidth=""1"" />
    </cols>
    <sheetData>
        <row r=""1"" spans=""1:2"" x14ac:dyDescent=""0.25"">
            <c r=""A1"" t=""str"">
                <v>Sheet1</v>
            </c>
        </row>
        <row r=""2"" spans=""1:2"" x14ac:dyDescent=""0.25"">
            <c r=""A2"" t=""str"">
                <v>FooCompany</v>
            </c>
            <c r=""B2"" s=""2"" />
        </row>
        <row r=""3"" spans=""1:2"" x14ac:dyDescent=""0.25"">
            <c r=""A3"" t=""str"">
                <v>Managers</v>
            </c>
        </row>
        <row r=""4"" spans=""1:2"" x14ac:dyDescent=""0.25"">
            <c r=""A4"" t=""str"">
                <v>{{managers.name}}</v>
            </c>
            <c r=""B4"" t=""str"">
                <v>{{managers.department}}</v>
            </c>
        </row>
        <row r=""5"" spans=""1:2"" x14ac:dyDescent=""0.25"">
            <c r=""A5"" t=""str"">
                <v>Employees</v>
            </c>
        </row>
        <row r=""6"" spans=""1:2"" x14ac:dyDescent=""0.25"">
            <c r=""A6"" t=""str"">
                <v>{{employees.name}}</v>
            </c>
            <c r=""B6"" t=""str"">
                <v>{{employees.department}}</v>
            </c>
        </row>
    </sheetData>
    <mergeCells count=""1"">
        <mergeCell ref=""A2:B2"" />
    </mergeCells>
    <pageMargins left=""0.78749999999999998"" right=""0.78749999999999998"" top=""1.05277777777778"" bottom=""1.05277777777778"" header=""0.78749999999999998"" footer=""0.78749999999999998"" />
    <pageSetup orientation=""portrait"" useFirstPageNumber=""1"" horizontalDpi=""300"" verticalDpi=""300"" />
    <headerFooter>
        <oddHeader>&amp;C&amp;""Times New Roman,Regular""&amp;12&amp;A</oddHeader>
        <oddFooter>&amp;C&amp;""Times New Roman,Regular""&amp;12Page &amp;P</oddFooter>
    </headerFooter>
</worksheet>";