<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>DocumentFormat.OpenXml</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>DocumentFormat.OpenXml.Packaging</Namespace>
  <Namespace>DocumentFormat.OpenXml.Spreadsheet</Namespace>
  <Namespace>DocumentFormat.OpenXml</Namespace>
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

//[How to: Parse and read a large spreadsheet document (Open XML SDK) | Microsoft Docs](https://docs.microsoft.com/en-us/office/open-xml/how-to-parse-and-read-a-large-spreadsheet)
void Main()
{
	var path = @"D:\git\MiniExcel\samples\xlsx\Test1,000,000x10\Test1,000,000x10.xlsx";
	ReadExcelFileSAX(path);
	
	//using (SpreadsheetDocument spreadsheetDocument =SpreadsheetDocument.Open(path, false))
	//{
	//	WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
	//	WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
	//	SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
	//	var firstRow = sheetData.Elements<Row>().First();
	//}
}


static void ReadExcelFileDOM(string fileName)
{
	using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
	{
		WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
		WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
		SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
		string text;
		foreach (Row r in sheetData.Elements<Row>())
		{
			foreach (Cell c in r.Elements<Cell>())
			{
				text = c.CellValue.Text;
				Console.Write(text + " ");
			}
		}
		Console.WriteLine();
		Console.ReadKey();
	}
}

// SAX模式
static void ReadExcelFileSAX(string fileName)
{
	using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
	{
		WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
		WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

		OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
		string text;
		reader.Read();
		if (reader.ElementType == typeof(CellValue))
		{
			text = reader.GetText();
			Console.Write(text + " ");
		}
	}
}