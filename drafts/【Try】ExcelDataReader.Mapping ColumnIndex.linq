<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>ExcelDataReader</NuGetReference>
  <NuGetReference>ExcelDataReader.Mapping</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>ExcelMapper</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.IO.Compression</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

//[hughbe/excel-mapper: An extension of ExcelDataReader that supports fluent mapping of rows to C# objects](https://github.com/hughbe/excel-mapper)

void Main()
{
	System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
	Foo();
}

// You can define other methods, fields, classes and namespaces here
public class Issue142VO
{
	public int MyProperty1 { get; set; }  //index = 1
	[ExcelIgnore]
	public int MyProperty7 { get; set; } //index = null.
	[ExcelColumnName("MyProperty2")]
	public int MyProperty9 { get; set; } //index = 3
	[ExcelColumnIndex(6)]
	public int MyProperty3 { get; set; } //index = 6
	[ExcelColumnIndex(0)] // equal column index 0
	public int MyProperty4 { get; set; }
	[ExcelColumnIndex(2)]
	public int MyProperty5 { get; set; } //index = 2
	public int MyProperty6 { get; set; } //index = 4
}

public void Foo()
{
	var excelFilePath = @"D:\git\MiniExcel\samples\xlsx\Issue142.xlsx";
	using var stream = File.OpenRead(excelFilePath);
	using var importer = new ExcelImporter(stream);

	ExcelSheet sheet = importer.ReadSheet();
	var events = sheet.ReadRows<Issue142VO>().ToArray();
	Console.WriteLine(events);
}