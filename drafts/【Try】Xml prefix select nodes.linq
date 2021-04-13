<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.IO.Compression</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

void Main()
{
	Execute(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<x:worksheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
	<x:sheetData>
        <x:row r=""1"">
            <x:c r=""A1"" t=""str"">
                <x:v>{{title}} {{date}}</x:v>
            </x:c>
        </x:row>		
    </x:sheetData>
</x:worksheet>");
	Execute(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
	<sheetData>
        <row r=""1"">
            <c r=""A1"" t=""str"">
                <x:v>{{title}} {{date}}</x:v>
            </c>
        </row>		
    </sheetData>
</worksheet>");
	Execute2(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
	<sheetData>
        <row r=""1"">
            <c r=""A1"" t=""str"">
                <v>{{title}} {{date}}</x:v>
            </c>
        </row>		
    </sheetData>
</worksheet>");
}

void Execute(string text)
{
	XmlDocument doc = new XmlDocument();
	{
		doc.LoadXml(text);

		XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
		ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		foreach (System.Xml.XmlElement r in doc.SelectNodes("/x:worksheet/x:sheetData/x:row", ns))
		{
			Console.WriteLine(r);
		}
	}
}

void Execute2(string text)
{
	XmlDocument doc = new XmlDocument();
	{
		doc.LoadXml(text);

		XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
		ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
		foreach (System.Xml.XmlElement r in doc.SelectNodes("/worksheet/sheetData/row", ns))
		{
			Console.WriteLine(r);
		}
	}
}

// You can define other methods, fields, classes and namespaces here
