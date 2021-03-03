<Query Kind="Program">
  <NuGetReference>AngleSharp</NuGetReference>
  <NuGetReference>ClosedXML</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>DocumentFormat.OpenXml</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <Namespace>ClosedXML.Excel</Namespace>
</Query>

void Main()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
	Console.WriteLine(path);
	using (var workbook = new XLWorkbook())
	{
		var worksheet = workbook.Worksheets.Add("Sample Sheet");
		worksheet.Cell("A1").Value = "\"<>+-*//}{\\n";
		worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
		worksheet.Cell("B1").Value = DateTime.Now;
		worksheet.Cell("B2").Value = 123;
		workbook.SaveAs(path);
	}
	Process.Start(new ProcessStartInfo() {FileName=path,UseShellExecute=true});
}

// You can define other methods, fields, classes and namespaces here
