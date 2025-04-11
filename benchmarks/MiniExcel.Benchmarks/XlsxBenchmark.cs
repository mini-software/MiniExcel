using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using ClosedXML.Report;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using MiniExcelLibs.Benchmarks.Utils;
using OfficeOpenXml;
using System.Text;

namespace MiniExcelLibs.Benchmarks;

public class XlsxBenchmark : BenchmarkBase
{
    [GlobalSetup]
    public void SetUp()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    [Benchmark(Description = "MiniExcel Template Generate")]
    public void MiniExcel_Template_Generate_Test()
    {
        const string templatePath = "TestTemplateBasicIEmumerableFill.xlsx";
        
        using var path = AutoDeletingPath.Create();
        var value = new
        {
            employees = Enumerable.Range(1, rowCount)
                .Select(s => new
                {
                    name = "Jack",
                    department = "HR"
                })
        };

        MiniExcel.SaveAsByTemplate(path.FilePath, templatePath, value);
    }

    [Benchmark(Description = "ClosedXml.Report Template Generate")]
    public void ClosedXml_Report_Template_Generate_Test()
    {
        const string templatePath = "TestTemplateBasicIEmumerableFill_ClosedXML_Report.xlsx";
        
        using var path = AutoDeletingPath.Create();
        var value = new
        {
            employees = Enumerable.Range(1, rowCount)
                .Select(s => new
                {
                    name = "Jack",
                    department = "HR"
                })
        };

        var template = new XLTemplate(templatePath);
        template.AddVariable(value);
        template.Generate();

        template.SaveAs(path.FilePath);
    }

    [Benchmark(Description = "MiniExcel QueryFirst")]
    public void MiniExcel_QueryFirst_Test()
    {
        _ = MiniExcel.Query(filePath).First();
    }

    [Benchmark(Description = "MiniExcel Query")]
    public void MiniExcel_Query()
    {
        foreach (var item in MiniExcel.Query(filePath))
        {
        }
    }

    [Benchmark(Description = "ExcelDataReader QueryFirst")]
    public void ExcelDataReader_QueryFirst_Test()
    {
        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        List<object> d = [];
        reader.Read();

        for (int i = 0; i < reader.FieldCount; i++)
            d.Add(reader.GetValue(i));
    }

    [Benchmark(Description = "ExcelDataReader Query")]
    public void ExcelDataReader_Query_Test()
    {
        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        while (reader.Read())
        {
            List<object> d = [];
            for (int i = 0; i < reader.FieldCount; i++)
                d.Add(reader.GetValue(i));
        }
    }

    [Benchmark(Description = "Epplus QueryFirst")]
    public void Epplus_QueryFirst_Test()
    {
        using var p = new ExcelPackage(new FileInfo(filePath));
        p.Workbook.Worksheets[0].Row(1);
    }

    [Benchmark(Description = "Epplus Query")]
    public void Epplus_Query_Test()
    {
        // [How do I iterate through rows in an excel table using epplus? - Stack Overflow] (https://stackoverflow.com/questions/21742038/how-do-i-iterate-through-rows-in-an-excel-table-using-epplus)

        using var p = new ExcelPackage(new FileInfo(filePath));

        var workSheet = p.Workbook.Worksheets[0];
        var start = workSheet.Dimension.Start;
        var end = workSheet.Dimension.End;

        for (int row = start.Row; row <= end.Row; row++)
        {
            for (int col = start.Column; col <= end.Column; col++)
            {
                object cellValue = workSheet.Cells[row, col].Text;
            }
        }
    }

    [Benchmark(Description = "ClosedXml QueryFirst")]
    public void ClosedXml_QueryFirst_Test()
    {
        using var workbook = new XLWorkbook(filePath);
        workbook.Worksheet(1).Row(1);
    }

    [Benchmark(Description = "ClosedXml Query")]
    public void ClosedXml_Query_Test()
    {
        using var workbook = new XLWorkbook(filePath);
        workbook.Worksheet(1).Rows();
    }

    [Benchmark(Description = "OpenXmlSDK QueryFirst")]
    public void OpenXmlSDK_QueryFirst_Test()
    {
        using var spreadsheetDocument = SpreadsheetDocument.Open(filePath, false);

        var workbookPart = spreadsheetDocument.WorkbookPart;
        var worksheetPart = workbookPart!.WorksheetParts.First();

        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
        var firstRow = sheetData.Elements<Row>().First();
    }

    [Benchmark(Description = "OpenXmlSDK Query")]
    public void OpenXmlSDK_Query_Test()
    {
        using var spreadsheetDocument = SpreadsheetDocument.Open(filePath, false);

        var workbookPart = spreadsheetDocument.WorkbookPart;
        var worksheetPart = workbookPart!.WorksheetParts.First();

        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
        var firstRow = sheetData.Elements<Row>().ToList();
    }

    [Benchmark(Description = "MiniExcel Create Xlsx")]
    public void MiniExcelCreateTest()
    {
        using var path = AutoDeletingPath.Create();
        MiniExcel.SaveAs(path.FilePath, GetValue());
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
