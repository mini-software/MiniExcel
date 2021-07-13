using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using OfficeOpenXml;

namespace MiniExcelLibs.Benchmarks
{
    public class XlsxBenchmark : BenchmarkBase
    {
        [GlobalSetup]
        public void SetUp()
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }

        [Benchmark(Description = "MiniExcel Template Generate")]
        public void MiniExcel_Template_Generate_Test()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                const string templatePath = @"TestTemplateBasicIEmumerableFill.xlsx";
                var value = new
                {
                    employees = Enumerable.Range(1, rowCount).Select(s => new { name = "Jack", department = "HR" })
                };
                MiniExcel.SaveAsByTemplate(path, templatePath, value);
            }
        }

        [Benchmark(Description = "ClosedXml.Report Template Generate")]
        public void ClosedXml_Report_Template_Generate_Test()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            var templatePath = @"TestTemplateBasicIEmumerableFill_ClosedXML_Report.xlsx";
            var template = new ClosedXML.Report.XLTemplate(templatePath);
            var value = new
            {
                employees = Enumerable.Range(1, rowCount).Select(s => new { name = "Jack", department = "HR" })
            };
            template.AddVariable(value);
            template.Generate();
            template.SaveAs(path);
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
            using (var p = new ExcelPackage(new FileInfo(filePath)))
            {
                p.Workbook.Worksheets[0].Row(1);
            }
        }

        [Benchmark(Description = "Epplus Query")]
        public void Epplus_Query_Test()
        {
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
            var value = Getvalue();
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            MiniExcel.SaveAs(path, value);
            File.Delete(path);
        }

        [Benchmark(Description = "ClosedXml Create Xlsx")]
        public void ClosedXmlCreateTest()
        {
            var value = Getvalue();
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Inserting Data");
                ws.Cell(1, 1).InsertData(value);
                wb.SaveAs(path);
            }
            File.Delete(path);
        }


        [Benchmark(Description = "Epplus Create Xlsx")]
        public void EpplusCreateTest()
        {
            var value = Getvalue();
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            using (var excelFile = new ExcelPackage(new FileInfo(path)))
            {
                var worksheet = excelFile.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells["A1"].LoadFromCollection(Collection: value, PrintHeaders: true);
                excelFile.Save();
            }
            File.Delete(path);
        }

        [Benchmark(Description = "OpenXmlSdk Create xlsx by DOM mode")]
        public void OpenXmlSdkCreateByDomModeTest()
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
                foreach (var item in Getvalue())
                {
                    var row = new Row();
                    row.Append(new Cell() { CellValue = new CellValue(item.Column1), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Column2), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Column3), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Column4), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Column5), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Column6), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Column7), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Column8), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Column9), DataType = CellValues.String });
                    row.Append(new Cell() { CellValue = new CellValue(item.Column10), DataType = CellValues.String });
                    sheetData.AppendChild(row);
                }

                workbookpart.Workbook.Save();

            }
            File.Delete(path);
        }
    }
}
