using System.Text;
using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MiniExcelLib.Benchmarks.Utils;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;

namespace MiniExcelLib.Benchmarks.BenchmarkSections;

public class CreateExcelBenchmark : BenchmarkBase
{
    private OpenXmlExporter _exporter;
        
    [GlobalSetup]
    public void SetUp()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        _exporter = new OpenXmlExporter();
    }

    [Benchmark(Description = "MiniExcel Create Xlsx")]
    public void MiniExcelCreateTest()
    {
        using var path = AutoDeletingPath.Create();
        _exporter.ExportExcel(path.FilePath, GetValue());
    }

    [Benchmark(Description = "ClosedXml Create Xlsx")]
    public void ClosedXmlCreateTest()
    {
        using var path = AutoDeletingPath.Create();
        using var wb = new XLWorkbook();

        var ws = wb.Worksheets.Add("Inserting Data");
        ws.Cell(1, 1).InsertData(GetValue());

        wb.SaveAs(path.FilePath);
    }

    [Benchmark(Description = "Epplus Create Xlsx")]
    public void EpplusCreateTest()
    {
        using var path = AutoDeletingPath.Create();
        using var excelFile = new ExcelPackage(new FileInfo(path.FilePath));

        var worksheet = excelFile.Workbook.Worksheets.Add("Sheet1");
        worksheet.Cells["A1"].LoadFromCollection(Collection: GetValue(), PrintHeaders: true);

        excelFile.Save();
    }
    [Benchmark(Description = "NPOI Create Xlsx")]
    public void NPOICreateTest()
    {
        using var path = AutoDeletingPath.Create();
        using var wb= new XSSFWorkbook();
        var worksheet = wb.CreateSheet("Sheet1");

        int i = 0;
        foreach (var item in GetValue())
        {
            var row = worksheet.CreateRow(i);
            row.CreateCell(0).SetCellValue(item.Column1);
            row.CreateCell(1).SetCellValue(item.Column2);
            row.CreateCell(2).SetCellValue(item.Column3);
            row.CreateCell(3).SetCellValue(item.Column4);
            row.CreateCell(4).SetCellValue(item.Column5);
            row.CreateCell(5).SetCellValue(item.Column6);
            row.CreateCell(6).SetCellValue(item.Column7);
            row.CreateCell(7).SetCellValue(item.Column8);
            row.CreateCell(8).SetCellValue(item.Column9);
            row.CreateCell(9).SetCellValue(item.Column10);
            i++;
        }

        using var fs = File.Create(path.FilePath);
        wb.Write(fs);
    }

    [Benchmark(Description = "OpenXmlSdk Create Xlsx by DOM mode")]
    public void OpenXmlSdkCreateByDomModeTest()
    {
        using var path = AutoDeletingPath.Create();
        using var spreadsheetDocument = SpreadsheetDocument.Create(path.FilePath, SpreadsheetDocumentType.Workbook);
        // By default, AutoSave = true, Editable = true, and Type = xlsx.

        WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
        workbookpart.Workbook = new Workbook();

        WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        Sheets sheets = spreadsheetDocument.WorkbookPart!.Workbook.AppendChild(new Sheets());

        sheets.Append(new Sheet
        {
            Id = spreadsheetDocument.WorkbookPart.
             GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1"
        });
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

        foreach (var item in GetValue())
        {
            sheetData!.AppendChild(new Row
            {
                item.Column1, item.Column2, item.Column3, item.Column4, item.Column5,
                item.Column6, item.Column7, item.Column8, item.Column9, item.Column10
            });
        }

        workbookpart.Workbook.Save();
    }
}
