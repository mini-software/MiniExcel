<Query Kind="Program">
  <NuGetReference>AngleSharp</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>DocumentFormat.OpenXml</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>DocumentFormat.OpenXml</Namespace>
  <Namespace>DocumentFormat.OpenXml.Packaging</Namespace>
  <Namespace>DocumentFormat.OpenXml.Spreadsheet</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.Globalization</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

void Main()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
	Console.WriteLine(path);

	OpenXmlSdkCreateTest();
}

public void OpenXmlSdkCreateTest()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
	Console.WriteLine(path);
	// Create a spreadsheet document by supplying the filepath.
	// By default, AutoSave = true, Editable = true, and Type = xlsx.
	SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
		 Create(path, SpreadsheetDocumentType.Workbook);

	// Add a WorkbookPart to the document.
	WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
	workbookpart.Workbook = new Workbook();

	// Add a WorksheetPart to the WorkbookPart.
	WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
	worksheetPart.Worksheet = new Worksheet(new SheetData());

	// Add Sheets to the Workbook.
	Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
		 AppendChild<Sheets>(new Sheets());

	// Append a new worksheet and associate it with the workbook.
	Sheet sheet = new Sheet()
	{
		Id = spreadsheetDocument.WorkbookPart.
		 GetIdOfPart(worksheetPart),
		SheetId = 1,
		Name = "Sheet1"
	};
	sheets.Append(sheet);

	var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
	foreach (var item in GetValues())
	{
		var row = new Row();
		row.Append(new Cell() { CellValue = new CellValue(item.text), DataType = CellValues.String });
		sheetData.AppendChild(row);
	}

	workbookpart.Workbook.Save();

	// Close the document.
	spreadsheetDocument.Close();
}

private static IEnumerable<dynamic> GetValues()
{
#if !DEBUG
	return Enumerable.Range(1, 10).Select(s => new { text = "Hello World" });
#else
            return Enumerable.Range(1, 1000000).Select(s => new { text = "Hello World" });
#endif
}