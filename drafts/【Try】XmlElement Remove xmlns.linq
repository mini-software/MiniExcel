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
	var doc = new XmlDocument();
	doc.LoadXml(xml);
	
	XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
	ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
	var dimension = doc.SelectSingleNode("/x:worksheet/x:dimension",ns) as XmlElement ;
	
	ns.RemoveNamespace("x","http://schemas.openxmlformats.org/spreadsheetml/2006/main");
	Console.WriteLine(dimension.OuterXml); //<dimension ref="A1:B6" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />
	
	dimension.RemoveAttributeNode("xmlns",dimension.NamespaceURI);
	Console.WriteLine(dimension.OuterXml); //<dimension ref="A1:B6" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />

	dimension.RemoveAttribute("xmlns");
	Console.WriteLine(dimension.OuterXml); //<dimension ref="A1:B6" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />
}

private static XElement RemoveAllNamespaces(XElement xmlDocument)
{
	if (!xmlDocument.HasElements)
	{
		XElement xElement = new XElement(xmlDocument.Name.LocalName);
		xElement.Value = xmlDocument.Value;

		foreach (XAttribute attribute in xmlDocument.Attributes())
			xElement.Add(attribute);

		return xElement;
	}
	return new XElement(xmlDocument.Name.LocalName, xmlDocument.Elements().Select(el => RemoveAllNamespaces(el)));
}

// You can define other methods, fields, classes and namespaces here
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