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

	CreateSpreadsheetWorkbook(path);
}

//[How to: Create a spreadsheet document by providing a file name (Open XML SDK) | Microsoft Docs](https://docs.microsoft.com/en-us/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)

// You can define other methods, fields, classes and namespaces here
public static void CreateSpreadsheetWorkbook(string filepath)
{
	// Create a spreadsheet document by supplying the filepath.
	// By default, AutoSave = true, Editable = true, and Type = xlsx.
	SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
		Create(filepath, SpreadsheetDocumentType.Workbook);

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
	for (int i = 0; i < 2; i++)
	{
		var row = new Row();
		row.Append(
			new Cell() { CellValue = new CellValue(1.ToString()), DataType = CellValues.Number }
			, new Cell() { CellValue = new CellValue(@"""<>+}{\nHello World"), DataType = CellValues.String }
			, new Cell() { CellValue = new CellValue("1"), DataType = CellValues.Boolean }
			, new Cell() { CellValue = new CellValue(DateTime.Now.ToString("s")), DataType = CellValues.Date }
		);
		sheetData.AppendChild(row);
	}

	workbookpart.Workbook.Save();

	// Close the document.
	spreadsheetDocument.Close();
}

// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
// If the cell already exists, returns it. 
private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
{
	Worksheet worksheet = worksheetPart.Worksheet;
	SheetData sheetData = worksheet.GetFirstChild<SheetData>();
	string cellReference = columnName + rowIndex;

	// If the worksheet does not contain a row with the specified row index, insert one.
	Row row;
	if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
	{
		row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
	}
	else
	{
		row = new Row() { RowIndex = rowIndex };
		sheetData.Append(row);
	}

	// If there is not a cell with the specified column name, insert one.  
	if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
	{
		return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
	}
	else
	{
		// Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
		Cell refCell = null;
		foreach (Cell cell in row.Elements<Cell>())
		{
			if (cell.CellReference.Value.Length == cellReference.Length)
			{
				if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
				{
					refCell = cell;
					break;
				}
			}
		}

		Cell newCell = new Cell() { CellReference = cellReference };
		row.InsertBefore(newCell, refCell);

		worksheet.Save();
		return newCell;
	}
}