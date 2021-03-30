using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Exporters;
using BenchmarkDotNet.Exporters.Csv;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Loggers;
using BenchmarkDotNet.Running;
using ClosedXML.Excel;
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
            var summary = BenchmarkRunner.Run<Benchmark>();
#else
            BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args, new DebugInProcessConfig());
#endif
            Console.Read();
        }
    }

    [BenchmarkCategory("Framework")]
    [MemoryDiagnoser]
    [SimpleJob(launchCount: 3, warmupCount: 3, targetCount: 3,invocationCount:3,baseline:false)]
    public abstract class BenchmarkBase
    {
        //public const string largeFilePath = @"D:\git\MiniExcel\samples\xlsx\Test1,000,000x10\Test1,000,000x10.xlsx";

        public const string filePath = @"D:\git\MiniExcel\samples\xlsx\Test10x10.xlsx";
    }

    public class Benchmark: BenchmarkBase
    {
        [GlobalSetup]
        public void SetUp()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        [Benchmark(Description = "MiniExcel QueryFirst")]
        public void MiniExcel_QueryFirst_Test()
        {
            MiniExcel.Query(filePath).First();
        }

        [Benchmark(Description = "ExcelDataReader QueryFirst")]
        public void ExcelDataReader_QueryFirst_Test()
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var d = new List<object>();
                reader.Read();
                for (int i = 0; i < reader.FieldCount; i++)
                    d.Add(reader.GetValue(i));
            }
        }

        [Benchmark(Description = "Epplus QueryFirst")]
        public void Epplus_QueryFirst_Test()
        {
            using (var p = new ExcelPackage(new FileInfo(filePath)))
            {
                p.Workbook.Worksheets[0].Row(1);
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
    }

    public class CreateTest
    {
        [Benchmark]
        public void MiniExcelCreateTest()
        {
            var values = Enumerable.Range(1, 10).Select((s, index) => new { index, value = Guid.NewGuid() });
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            using (var stream = File.Create(path))
                stream.SaveAs(values);
        }
    }
}
