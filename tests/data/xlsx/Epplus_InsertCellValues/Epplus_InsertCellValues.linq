<Query Kind="Program">
  <NuGetReference>AngleSharp</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>DocumentFormat.OpenXml</NuGetReference>
  <NuGetReference Version="4.5.3">EPPlus</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <Namespace>OfficeOpenXml</Namespace>
</Query>

void Main()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
	Console.WriteLine(path);
	using (var p = new ExcelPackage())
	{
		//A workbook must have at least on cell, so lets add one... 
		var ws = p.Workbook.Worksheets.Add("Sheet1");
		//To set values in the spreadsheet use the Cells indexer.

		ws.Cells["A1"].Value = "\"<>+-*//}{\\n";
		ws.Cells["A2"].Formula = "=MID(A1, 7, 5)";
		ws.Cells["B1"].Value = DateTime.Now;
		ws.Cells["B2"].Value = 123;
		//Save the new workbook. We haven't specified the filename so use the Save as method.
		p.SaveAs(new FileInfo(path));
	}

	Process.Start(new ProcessStartInfo() {FileName=path,UseShellExecute=true});
}
