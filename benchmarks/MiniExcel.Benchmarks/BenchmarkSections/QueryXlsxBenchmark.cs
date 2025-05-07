using System.Text;
using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using OfficeOpenXml;

namespace MiniExcelLibs.Benchmarks.BenchmarkSections;

public class QueryXlsxBenchmark : BenchmarkBase
{
    [GlobalSetup]
    public void SetUp()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    [Benchmark(Description = "MiniExcel QueryFirst")]
    public void MiniExcel_QueryFirst_Test()
    {
        _ = MiniExcel.Query(FilePath).First();
    }

    [Benchmark(Description = "MiniExcel Query")]
    public void MiniExcel_Query()
    {
        foreach (var _ in MiniExcel.Query(FilePath)) { }
    }

    [Benchmark(Description = "ExcelDataReader QueryFirst")]
    public void ExcelDataReader_QueryFirst_Test()
    {
        using var stream = File.Open(FilePath, FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);
    
        List<object> d = [];
        reader.Read();
    
        for (var i = 0; i < reader.FieldCount; i++)
            d.Add(reader.GetValue(i));
    }
    
    [Benchmark(Description = "ExcelDataReader Query")]
    public void ExcelDataReader_Query_Test()
    {
        using var stream = File.Open(FilePath, FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);
    
        while (reader.Read())
        {
            List<object> d = [];
            for (var i = 0; i < reader.FieldCount; i++)
                d.Add(reader.GetValue(i));
        }
    }
    
    [Benchmark(Description = "Epplus QueryFirst")]
    public void Epplus_QueryFirst_Test()
    {
        using var p = new ExcelPackage(new FileInfo(FilePath));
        p.Workbook.Worksheets[0].Row(1);
    }
    
    [Benchmark(Description = "Epplus Query")]
    public void Epplus_Query_Test()
    {
        // [How do I iterate through rows in an excel table using epplus? - Stack Overflow] (https://stackoverflow.com/questions/21742038/how-do-i-iterate-through-rows-in-an-excel-table-using-epplus)
    
        using var p = new ExcelPackage(new FileInfo(FilePath));
    
        var workSheet = p.Workbook.Worksheets[0];
        var start = workSheet.Dimension.Start;
        var end = workSheet.Dimension.End;
    
        for (var row = start.Row; row <= end.Row; row++)
        {
            for (var col = start.Column; col <= end.Column; col++)
            {
                object cellValue = workSheet.Cells[row, col].Text;
            }
        }
    }
    
    [Benchmark(Description = "ClosedXml QueryFirst")]
    public void ClosedXml_QueryFirst_Test()
    {
        using var workbook = new XLWorkbook(FilePath);
        workbook.Worksheet(1).Row(1);
    }
    
    [Benchmark(Description = "ClosedXml Query")]
    public void ClosedXml_Query_Test()
    {
        using var workbook = new XLWorkbook(FilePath);
        workbook.Worksheet(1).Rows();
    }
    
    [Benchmark(Description = "OpenXmlSDK QueryFirst")]
    public void OpenXmlSDK_QueryFirst_Test()
    {
        using var spreadsheetDocument = SpreadsheetDocument.Open(FilePath, false);
    
        var workbookPart = spreadsheetDocument.WorkbookPart;
        var worksheetPart = workbookPart!.WorksheetParts.First();
    
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
        var firstRow = sheetData.Elements<Row>().First();
    }
    
    [Benchmark(Description = "OpenXmlSDK Query")]
    public void OpenXmlSDK_Query_Test()
    {
        using var spreadsheetDocument = SpreadsheetDocument.Open(FilePath, false);
    
        var workbookPart = spreadsheetDocument.WorkbookPart;
        var worksheetPart = workbookPart!.WorksheetParts.First();
    
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
        var firstRow = sheetData.Elements<Row>().ToList();
    }
}