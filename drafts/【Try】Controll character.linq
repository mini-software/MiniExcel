<Query Kind="Program">
  <NuGetReference>ClosedXML</NuGetReference>
  <NuGetReference>DocumentFormat.OpenXml</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <Namespace>ClosedXML.Excel</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.IO.Compression</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <Namespace>DocumentFormat.OpenXml.Packaging</Namespace>
  <Namespace>DocumentFormat.OpenXml.Spreadsheet</Namespace>
  <Namespace>DocumentFormat.OpenXml</Namespace>
</Query>

void Main()
{
	//Console.WriteLine(chars);
	Text();
	//OpenXml();
	//ClosedXml();
}

void Text()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.txt");
	var sb = new StringBuilder();
	foreach (var c in chars)
	{
		sb.AppendLine(c.ToString() + "  k ");
	}
	File.AppendAllText(path, sb.ToString());
	Console.WriteLine(path);
}

void OpenXml()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
	{
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
		{

			foreach (var c in chars)
			{
				var row = new Row();
				//row.Append(new Cell() { CellValue = new CellValue("D" + c), DataType = CellValues.String });
				row.Append(new Cell() { CellValue = new CellValue("E" + c.ToString()), DataType = CellValues.String });
				sheetData.AppendChild(row);
			}


		}

		workbookpart.Workbook.Save();

		// Close the document.
		spreadsheetDocument.Close();
	}
	Console.WriteLine(path);
}

void ClosedXml()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
	using (var workbook = new XLWorkbook())
	{
		var worksheet = workbook.Worksheets.Add("Sample Sheet");

		var index = 0;
		foreach (var c in chars)
		{
			index++;
			worksheet.Cell($"A{index}").Value = c;
		}

		workbook.SaveAs(path);
	}
	Console.WriteLine(path);
}

void MiniExcel()
{
	var input = Enumerable.Range(1, 10).Select(s => new { Test1 = '\u0006', Test2 = "123" });
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

	MiniExcelLibs.MiniExcel.SaveAs(path, input);

	Console.WriteLine(path);
}

// You can define other methods, fields, classes and namespaces here
static char[] chars = new char[] {
	'\u0000',
	'\u0001',
	'\u0002',
	'\u0003',
	'\u0004',
	'\u0005',
	'\u0006',
	'\u0007',
	'\u0008',
	'\u0009', //<HT>
	'\u000A', //<LF>
	'\u000B',
	'\u000C',
	'\u000D', //<CR>
	'\u000E',
	'\u000F',
	'\u0010',
	'\u0011',
	'\u0012',
	'\u0013',
	'\u0014',
	'\u0015',
	'\u0016',
	'\u0017',
	'\u0018',
	'\u0019',
	'\u001A',
	'\u001B',
	'\u001C',
	'\u001D',
	'\u001E',
	'\u001F',
	'\u007F'
};