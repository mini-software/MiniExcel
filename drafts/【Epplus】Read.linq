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
	Stopwatch sw = new Stopwatch();
	sw.Start();
	Console.WriteLine("start memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB");
	
	var path = @"D:\git\MiniExcel\samples\xlsx\Test1,000,000x10\Test1,000,000x10.xlsx";
	using (var p = new ExcelPackage(new FileInfo(path)))
	{
		Console.WriteLine("A1 : " + p.Workbook.Worksheets[0].Cells["A1"].Value);
		Console.WriteLine("end memory usage: " +System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB & run time : {sw.ElapsedMilliseconds}ms");
	}
}
