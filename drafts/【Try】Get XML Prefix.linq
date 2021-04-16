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

// [xml - C#: How to get the name (with prefix) from XElement as string? - Stack Overflow](https://stackoverflow.com/questions/6387726/c-how-to-get-the-name-with-prefix-from-xelement-as-string)
void Main()
{
	{
		var doc = XElement.Parse(xmlWithPrefix);
		var prefix = doc.GetPrefixOfNamespace(doc.Name.Namespace);
		Console.WriteLine(prefix);
	}
	{
		var doc = XElement.Parse(XmlWithoutPrefix);
		var prefix = doc.GetPrefixOfNamespace(doc.Name.Namespace);
		Console.WriteLine(prefix);
	}
	{
		var doc = new XmlDocument();
		doc.LoadXml(xmlWithPrefix);
		Console.WriteLine(doc.ChildNodes[1].Prefix);
		doc.ChildNodes[1].Prefix = "";
	}
}

// You can define other methods, fields, classes and namespaces here
const string xmlWithPrefix = @"<?xml version=""1.0"" encoding=""utf-8""?>
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

const string XmlWithoutPrefix = @"<?xml version=""1.0"" encoding=""utf-8""?>
<worksheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <sheetPr>
    <outlinePr summaryBelow=""1"" summaryRight=""1"" />
  </sheetPr>
  <dimension ref=""A1:B2"" />
  <sheetViews>
    <sheetView workbookViewId=""0"" />
  </sheetViews>
  <sheetFormatPr defaultRowHeight=""15"" />
  <sheetData>
    <row r=""1"" spans=""1:2"">
      <c r=""A1"" s=""0"" t=""s"">
        <v>0</v>
      </c>
      <c r=""B1"" s=""1"">
        <v>44257.3802667361</v>
      </c>
    </row>
    <row r=""2"" spans=""1:2"">
      <c r=""A2"" s=""0"">
        <f>MID(A1, 7, 5)</f>
      </c>
      <c r=""B2"" s=""0"" t=""n"">
        <v>123</v>
      </c>
    </row>
  </sheetData>
  <printOptions horizontalCentered=""0"" verticalCentered=""0"" headings=""0"" gridLines=""0"" />
  <pageMargins left=""0.75"" right=""0.75"" top=""0.75"" bottom=""0.5"" header=""0.5"" footer=""0.75"" />
  <pageSetup paperSize=""1"" scale=""100"" pageOrder=""downThenOver"" orientation=""default"" blackAndWhite=""0"" draft=""0"" cellComments=""none"" errors=""displayed"" />
  <headerFooter />
  <tableParts count=""0"" />
</worksheet>";