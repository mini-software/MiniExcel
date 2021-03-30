<Query Kind="Program">
  <NuGetReference>AngleSharp</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>DocumentFormat.OpenXml</NuGetReference>
  <NuGetReference>EPPlus</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>OfficeOpenXml</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <Namespace>DocumentFormat.OpenXml.Spreadsheet</Namespace>
</Query>

void Main()
{
	// If you use EPPlus in a noncommercial context
	// according to the Polyform Noncommercial license:
	ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
	
	//IEnumerable<object> collection =  Enumerable.Range(1, 10).Select(s => new Demo{ text = "Hello World" }); //IEnumerable<object> will be null
	var collection =  Enumerable.Range(1, 10).Select(s => new Demo{ text = "Hello World" });
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
	using (var excelFile = new ExcelPackage(new FileInfo(path)))
	{
		var worksheet = excelFile.Workbook.Worksheets.Add("Sheet1");
		worksheet.Cells["A1"].LoadFromCollection(collection, true);
		excelFile.Save();
	}
	Console.WriteLine(path);
	//File.Delete(path);
}

public class Demo
{
	public string text { get; set; }
}