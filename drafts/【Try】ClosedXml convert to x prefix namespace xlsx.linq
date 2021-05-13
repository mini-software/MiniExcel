<Query Kind="Program">
  <NuGetReference>ClosedXML</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
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
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
	Console.WriteLine(path);
	using (var workbook = new XLWorkbook(@"D:\git\MiniExcel\samples\xlsx\TestIssue207_Template_Merge_row_list_rendering_without_merge\template.xlsx"))
	{
		workbook.SaveAs(path);
	}
}

// You can define other methods, fields, classes and namespaces here
