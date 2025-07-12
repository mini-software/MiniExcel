using MiniExcelLib.OpenXml.Attributes;
using MiniExcelLib.OpenXml.Models;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.Tests;

public class MiniExcelOpenXmlMultipleSheetTests
{
    private readonly OpenXmlImporter _importer =  MiniExcel.GetImporterProvider().GetExcelImporter();
    
    [Fact]
    public void SpecifySheetNameQueryTest()
    {
        const string path = "../../../../../samples/xlsx/TestMultiSheet.xlsx";
        {
            var rows =  _importer.QueryExcel(path, sheetName: "Sheet3").ToList();
            Assert.Equal(5, rows.Count);
            Assert.Equal(3, rows[0].A);
            Assert.Equal(3, rows[0].B);
        }
        {
            var rows =  _importer.QueryExcel(path, sheetName: "Sheet2").ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(1, rows[0].A);
            Assert.Equal(1, rows[0].B);
        }
        {
            var rows =  _importer.QueryExcel(path, sheetName: "Sheet1").ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(2, rows[0].A);
            Assert.Equal(2, rows[0].B);
        }
        Assert.Throws<InvalidOperationException>(() =>  _importer.QueryExcel(path, sheetName: "xxxx").ToList());

        using var stream = File.OpenRead(path);
        
        {
            var rows =  _importer.QueryExcel(stream, sheetName: "Sheet3").ToList();
            Assert.Equal(5, rows.Count);
            Assert.Equal(3, rows[0].A);
            Assert.Equal(3, rows[0].B);
        }
        {
            var rows =  _importer.QueryExcel(stream, sheetName: "Sheet2").ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(1, rows[0].A);
            Assert.Equal(1, rows[0].B);
        }
        {
            var rows =  _importer.QueryExcel(stream, sheetName: "Sheet1").ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(2, rows[0].A);
            Assert.Equal(2, rows[0].B);
        }
        {
            var rows =  _importer.QueryExcel(stream, sheetName: "Sheet1").ToList();
            Assert.Equal(12, rows.Count);
            Assert.Equal(2, rows[0].A);
            Assert.Equal(2, rows[0].B);
        }
    }

    [Fact]
    public void MultiSheetsQueryBasicTest()
    {
        const string path = "../../../../../samples/xlsx/TestMultiSheet.xlsx";
        using var stream = File.OpenRead(path);
        
        _ =  _importer.QueryExcel(stream, sheetName: "Sheet1");
        _ =  _importer.QueryExcel(stream, sheetName: "Sheet2");
        _ =  _importer.QueryExcel(stream, sheetName: "Sheet3");
    }

    [Fact]
    public void MultiSheetsQueryTest()
    {
        const string path = "../../../../../samples/xlsx/TestMultiSheet.xlsx";
        {
            var sheetNames =  _importer.GetSheetNames(path).ToList();
            foreach (var sheetName in sheetNames)
            {
                var rows =  _importer.QueryExcel(path, sheetName: sheetName).ToList();
            }

            Assert.Equal(new[] { "Sheet1", "Sheet2", "Sheet3" }, sheetNames);
        }

        {
            using var stream = File.OpenRead(path);
            var sheetNames =  _importer.GetSheetNames(stream).ToList();
            Assert.Equal(new[] { "Sheet1", "Sheet2", "Sheet3" }, sheetNames);
            
            foreach (var sheetName in sheetNames)
            {
                var rows =  _importer.QueryExcel(stream, sheetName: sheetName).ToList();
            }
        }
    }

    [ExcelSheet(Name = "Users")]
    private class UserDto
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    [ExcelSheet(Name = "Departments", State = SheetState.Hidden)]
    private class DepartmentDto
    {
        public string ID { get; set; }
        public string Name { get; set; }
    }

    [Fact]
    public void ExcelSheetAttributeIsUsedWhenReadExcel()
    {
        const string path = "../../../../../samples/xlsx/TestDynamicSheet.xlsx";
        using (var stream = File.OpenRead(path))
        {
            var users =  _importer.QueryExcel<UserDto>(stream).ToList();
            Assert.Equal(2, users.Count);
            Assert.Equal("Jack", users[0].Name);

            var departments =  _importer.QueryExcel<DepartmentDto>(stream).ToList();
            Assert.Equal(2, departments.Count);
            Assert.Equal("HR", departments[0].Name);
        }

        {
            var users =  _importer.QueryExcel<UserDto>(path).ToList();
            Assert.Equal(2, users.Count);
            Assert.Equal("Jack", users[0].Name);

            var departments =  _importer.QueryExcel<DepartmentDto>(path).ToList();
            Assert.Equal(2, departments.Count);
            Assert.Equal("HR", departments[0].Name);
        }
    }

    [Fact]
    public void DynamicSheetConfigurationIsUsedWhenReadExcel()
    {
        var configuration = new OpenXmlConfiguration
        {
            DynamicSheets =
            [
                new DynamicExcelSheet("usersSheet") { Name = "Users" },
                new DynamicExcelSheet("departmentSheet") { Name = "Departments" }
            ]
        };

        const string path = "../../../../../samples/xlsx/TestDynamicSheet.xlsx";
        using (var stream = File.OpenRead(path))
        {
            // take first sheet as default
            var users =  _importer.QueryExcel(stream, configuration: configuration, useHeaderRow: true).ToList();
            Assert.Equal(2, users.Count);
            Assert.Equal("Jack", users[0].Name);

            // take second sheet by sheet name
            var departments =  _importer.QueryExcel(stream, sheetName: "Departments", configuration: configuration, useHeaderRow: true).ToList();
            Assert.Equal(2, departments.Count);
            Assert.Equal("HR", departments[0].Name);

            // take second sheet by sheet key
            departments =  _importer.QueryExcel(stream, sheetName: "departmentSheet", configuration: configuration, useHeaderRow: true).ToList();
            Assert.Equal(2, departments.Count);
            Assert.Equal("HR", departments[0].Name);
        }

        {
            // take first sheet as default
            var users =  _importer.QueryExcel(path, configuration: configuration, useHeaderRow: true).ToList();
            Assert.Equal(2, users.Count);
            Assert.Equal("Jack", users[0].Name);

            // take second sheet by sheet name
            var departments =  _importer.QueryExcel(path, sheetName: "Departments", configuration: configuration, useHeaderRow: true).ToList();
            Assert.Equal(2, departments.Count);
            Assert.Equal("HR", departments[0].Name);

            // take second sheet by sheet key
            departments =  _importer.QueryExcel(path, sheetName: "departmentSheet", configuration: configuration, useHeaderRow: true).ToList();
            Assert.Equal(2, departments.Count);
            Assert.Equal("HR", departments[0].Name);
        }
    }

    [Fact]
    public void ReadSheetVisibilityStateTest()
    {
        const string path = "../../../../../samples/xlsx/TestMultiSheetWithHiddenSheet.xlsx";
        {
            var sheetInfos =  _importer.GetSheetInformations(path).ToList();
            Assert.Collection(sheetInfos,
                i =>
                {
                    Assert.Equal(0u, i.Index);
                    Assert.Equal(2u, i.Id);
                    Assert.Equal(SheetState.Visible, i.State);
                    Assert.Equal("Sheet2", i.Name);
                },
                i =>
                {
                    Assert.Equal(1u, i.Index);
                    Assert.Equal(1u, i.Id);
                    Assert.Equal(SheetState.Visible, i.State);
                    Assert.Equal("Sheet1", i.Name);
                },
                i =>
                {
                    Assert.Equal(2u, i.Index);
                    Assert.Equal(3u, i.Id);
                    Assert.Equal(SheetState.Visible, i.State);
                    Assert.Equal("Sheet3", i.Name);
                },
                i =>
                {
                    Assert.Equal(3u, i.Index);
                    Assert.Equal(5u, i.Id);
                    Assert.Equal(SheetState.Hidden, i.State);
                    Assert.Equal("HiddenSheet4", i.Name);
                });
        }
    }

    [Fact]
    public void WriteHiddenSheetTest()
    {
        var configuration = new OpenXmlConfiguration
        {
            DynamicSheets =
            [
                new DynamicExcelSheet("usersSheet")
                {
                    Name = "Users",
                    State = SheetState.Visible
                },
                new DynamicExcelSheet("departmentSheet")
                {
                    Name = "Departments",
                    State = SheetState.Hidden
                }
            ]
        };

        var users = new[] { new { Name = "Jack", Age = 25 }, new { Name = "Mike", Age = 44 } };
        var department = new[] { new { ID = "01", Name = "HR" }, new { ID = "02", Name = "IT" } };
        var sheets = new Dictionary<string, object>
        {
            ["usersSheet"] = users,
            ["departmentSheet"] = department
        };

        using var file = AutoDeletingPath.Create();
        var path = file.ToString();

        var rowsWritten =  MiniExcel.GetExporterProvider().GetExcelExporter().ExportExcel(path, sheets, configuration: configuration);
        Assert.Equal(2, rowsWritten.Length);
        Assert.Equal(2, rowsWritten[0]);


        var sheetInfos =  _importer.GetSheetInformations(path).ToList();
        Assert.Collection(sheetInfos,
            i =>
            {
                Assert.Equal(0u, i.Index);
                Assert.Equal(1u, i.Id);
                Assert.Equal(SheetState.Visible, i.State);
                Assert.Equal("Users", i.Name);
            },
            i =>
            {
                Assert.Equal(1u, i.Index);
                Assert.Equal(2u, i.Id);
                Assert.Equal(SheetState.Hidden, i.State);
                Assert.Equal("Departments", i.Name);
            });

        foreach (var sheetName in sheetInfos.Select(s => s.Name))
        {
            var rows =  _importer.QueryExcel(path, sheetName: sheetName).ToList();
        }
    }
}