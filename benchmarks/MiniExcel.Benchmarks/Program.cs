using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Order;
using BenchmarkDotNet.Running;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using MiniExcelLibs;
using OfficeOpenXml;

namespace MiniExcelLibs.Benchmarks
{
    class Program
    {
        static void Main(string[] args)
        {
#if !DEBUG
            //new BenchmarkSwitcher(typeof(Program).Assembly).Run(args, new Config());
            BenchmarkRunner.Run<XlsxBenchmark>();
#else
            //BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args, new DebugInProcessConfig());
            new XlsxBenchmark().ClosedXml_Query_Test();
#endif
            Console.Read();
        }
    }

    [BenchmarkCategory("Framework")]
    [MemoryDiagnoser]
    [SimpleJob(launchCount: 1, warmupCount: 1, targetCount: 1,invocationCount:1,baseline:false)]
    [Orderer(SummaryOrderPolicy.FastestToSlowest)]
    public abstract class BenchmarkBase
    {
#if !DEBUG
        public const string filePath = @"D:\git\MiniExcel\samples\xlsx\Test1,000,000x10\Test1,000,000x10.xlsx";
        //public const string filePath = @"D:\git\MiniExcel\samples\xlsx\Test10x10.xlsx";
#else
        public const string filePath = @"D:\git\MiniExcel\samples\xlsx\Test1,000,000x10\Test1,000,000x10.xlsx";
        //public const string filePath = @"D:\git\MiniExcel\samples\xlsx\Test10x10.xlsx";
#endif
    }

    public class XlsxBenchmark: BenchmarkBase
    {
        [GlobalSetup]
        public void SetUp()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        [Benchmark(Description = "MiniExcel QueryFirst")]
        public void MiniExcel_QueryFirst_Test()
        {
            MiniExcel.Query(filePath).First();
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
#if DEBUG
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
#endif
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var d = new List<object>();
                reader.Read();
                for (int i = 0; i < reader.FieldCount; i++)
                    d.Add(reader.GetValue(i));
            }
        }

        [Benchmark(Description = "ExcelDataReader Query")]
        public void ExcelDataReader_Query_Test()
        {
#if DEBUG
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
#endif
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                while (reader.Read())
                {
                    var d = new List<object>();
                    for (int i = 0; i < reader.FieldCount; i++)
                        d.Add(reader.GetValue(i));
                }
            }
        }

        [Benchmark(Description = "Epplus QueryFirst")]
        public void Epplus_QueryFirst_Test()
        {
#if DEBUG
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
#endif
            using (var p = new ExcelPackage(new FileInfo(filePath)))
            {
                p.Workbook.Worksheets[0].Row(1);
            }
        }

        [Benchmark(Description = "Epplus Query")]
        public void Epplus_Query_Test()
        {
#if DEBUG
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
#endif
            // [c# - How do I iterate through rows in an excel table using epplus? - Stack Overflow](https://stackoverflow.com/questions/21742038/how-do-i-iterate-through-rows-in-an-excel-table-using-epplus)
            using (var p = new ExcelPackage(new FileInfo(filePath)))
            {
                var workSheet = p.Workbook.Worksheets[0];
                var start = workSheet.Dimension.Start;
                var end = workSheet.Dimension.End;
                for (int row = start.Row; row <= end.Row; row++)
                { // Row by row...
                    for (int col = start.Column; col <= end.Column; col++)
                    { // ... Cell by cell...
                        object cellValue = workSheet.Cells[row, col].Text; // This got me the actual value I needed.
                    }
                }
            }
        }

        [Benchmark(Description = "ClosedXml QueryFirst")]
        public void ClosedXml_QueryFirst_Test()
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                workbook.Worksheet(1).Row(1);
            }
        }

        [Benchmark(Description = "ClosedXml Query")]
        public void ClosedXml_Query_Test()
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                workbook.Worksheet(1).Rows();
            }
        }

        [Benchmark(Description = "OpenXmlSDK QueryFirst")]
        public void OpenXmlSDK_QueryFirst_Test()
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                var firstRow = sheetData.Elements<Row>().First();
            }
        }

        [Benchmark(Description = "OpenXmlSDK Query")]
        public void OpenXmlSDK_Query_Test()
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                var firstRow = sheetData.Elements<Row>().ToList();
            }
        }

        [Benchmark(Description = "MiniExcel Create Xlsx")]
        public void MiniExcelCreateTest()
        {
            var values = GetValues();
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            using (var stream = File.Create(path))
                stream.SaveAs(values);
            File.Delete(path);
        }

        [Benchmark(Description = "ClosedXml Create Xlsx")]
        public void ClosedXmlCreateTest()
        {
            var values = GetValues();
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Inserting Data");
                ws.Cell(1, 1).InsertData(values);
                wb.SaveAs(path);
            }

            File.Delete(path);
        }


        [Benchmark(Description = "Epplus Create Xlsx")]
        public void EpplusCreateTest()
        {
#if DEBUG
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
#endif
            var values = GetValues();
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            using (var excelFile = new ExcelPackage(new FileInfo(path)))
            {
                var worksheet = excelFile.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells["A1"].LoadFromCollection(Collection: values, PrintHeaders: true);
                excelFile.Save();
            }
            File.Delete(path);
        }

        [Benchmark(Description = "OpenXmlSdk Create Xlsx")]
        public void OpenXmlSdkCreateTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
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
                    row.Append(new Cell() { CellValue = new CellValue(item.Text1), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Text2), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Text3), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Text4), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Text5), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Text6), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Text7), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Text8), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Text9), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Text10), DataType = CellValues.String });
                    sheetData.AppendChild(row);
                }

                workbookpart.Workbook.Save();

            }
            File.Delete(path);
        }

        private static IEnumerable<Demo> GetValues()
        {
#if !DEBUG
            return Enumerable.Range(1, 1000000).Select(s => new Demo());
#else
            return Enumerable.Range(1, 1000000).Select(s => new Demo());
#endif
        }

        public class Demo
        {
            public string Text1 { get; set; } = "Hello World";
            public string Text2 { get; set; } = "Hello World";
            public string Text3 { get; set; } = "Hello World";
            public string Text4 { get; set; } = "Hello World";
            public string Text5 { get; set; } = "Hello World";
            public string Text6 { get; set; } = "Hello World";
            public string Text7 { get; set; } = "Hello World";
            public string Text8 { get; set; } = "Hello World";
            public string Text9 { get; set; } = "Hello World";
            public string Text10 { get; set; } = "Hello World";
        }
    }
}
