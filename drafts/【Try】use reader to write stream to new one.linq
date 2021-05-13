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
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

// https://docs.microsoft.com/en-us/archive/msdn-magazine/2003/may/manipulating-xml-data-with-integrated-readers-and-writers-in-net
void Main()
{
	var bytes = Encoding.UTF8.GetBytes(xml);
	using (var stream = new MemoryStream(bytes))
	{
		GetXmlFileNodeLayout(stream).Dump();
	}
}

// You can define other methods, fields, classes and namespaces here
string GetXmlFileNodeLayout(System.IO.Stream stream)
{
	// Open the stream 
	XmlTextReader reader = new XmlTextReader(stream);
	// Loop through the nodes and accumulate text into a string 
	StringWriter writer = new StringWriter();
	string tabPrefix = "";
	while (reader.Read())
	{
		// Write the start tag 
		if (reader.NodeType == XmlNodeType.Element)
		{
			tabPrefix = new string('\t', reader.Depth);
			writer.WriteLine("{0}<{1}>", tabPrefix, reader.Name);
		}
		else
		{
			// Write the end tag 
			if (reader.NodeType == XmlNodeType.EndElement)
			{
				tabPrefix = new string('\t', reader.Depth);
				writer.WriteLine("{0}</{1}>", tabPrefix, reader.Name);
			}
		}
	}
	// Write to the output window 
	string buf = writer.ToString();
	writer.Close();
	// Close the stream 
	reader.Close();
	return buf;
}

const string xml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:worksheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:sheetPr>
        <x:outlinePr summaryBelow=""1"" summaryRight=""1"" />
    </x:sheetPr>
    <x:dimension ref=""A1:B2"" />
    <x:sheetViews>
        <x:sheetView workbookViewId=""0"" />
    </x:sheetViews>
    <x:sheetFormatPr defaultRowHeight=""15"" />
    <x:sheetData>
        <x:row r=""1"" spans=""1:2"">
            <x:c r=""A1"" s=""0"" t=""s"">
                <x:v>0</x:v>
            </x:c>
            <x:c r=""B1"" s=""1"">
                <x:v>44257.3802667361</x:v>
            </x:c>
        </x:row>
        <x:row r=""2"" spans=""1:2"">
            <x:c r=""A2"" s=""0"">
                <x:f>MID(A1, 7, 5)</x:f>
            </x:c>
            <x:c r=""B2"" s=""0"" t=""n"">
                <x:v>123</x:v>
            </x:c>
        </x:row>
    </x:sheetData>
    <x:printOptions horizontalCentered=""0"" verticalCentered=""0"" headings=""0"" gridLines=""0"" />
    <x:pageMargins left=""0.75"" right=""0.75"" top=""0.75"" bottom=""0.5"" header=""0.5"" footer=""0.75"" />
    <x:pageSetup paperSize=""1"" scale=""100"" pageOrder=""downThenOver"" orientation=""default"" blackAndWhite=""0"" draft=""0"" cellComments=""none"" errors=""displayed"" />
    <x:headerFooter />
    <x:tableParts count=""0"" />
</x:worksheet>";

