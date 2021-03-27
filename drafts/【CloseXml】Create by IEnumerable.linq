<Query Kind="Program">
  <NuGetReference>ClosedXML</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>ClosedXML.Excel</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Diagnostics</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Text</RemoveNamespace>
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

void Main()
{
	var values = Enumerable.Range(1, 1048575).Select((s, index) => new { index, value = Guid.NewGuid() }).ToList();
	using (var workbook = new XLWorkbook())
	{
		var ws = workbook.Worksheets.Add("Sample Sheet");
		ws.Cell(1, 1).InsertTable(values);
		workbook.SaveAs("HelloWorld.xlsx");
	}
}

// You can define other methods, fields, classes and namespaces here
