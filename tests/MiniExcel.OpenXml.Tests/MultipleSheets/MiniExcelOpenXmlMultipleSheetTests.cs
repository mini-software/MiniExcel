using MiniExcelLib.OpenXml.Models;
using MiniExcelLib.Tests.Common.Utils;

namespace MiniExcelLib.OpenXml.Tests.MultipleSheets;

public class MiniExcelOpenXmlMultipleSheetTests
{
    private readonly OpenXmlImporter _excelImporter =  MiniExcel.Importers.GetOpenXmlImporter();
    
    [Fact]
    public void SpecifySheetNameQueryTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        var rows1 = _excelImporter.Query(path, sheetName: "Sheet3").ToList();
        Assert.Equal(5, rows1.Count);
        Assert.Equal(3, rows1[0].A);
        Assert.Equal(3, rows1[0].B);

        var rows2 = _excelImporter.Query(path, sheetName: "Sheet2").ToList();
        Assert.Equal(12, rows2.Count);
        Assert.Equal(1, rows2[0].A);
        Assert.Equal(1, rows2[0].B);
        
        var rows3 = _excelImporter.Query(path, sheetName: "Sheet1").ToList();
        Assert.Equal(12, rows3.Count);
        Assert.Equal(2, rows3[0].A);
        Assert.Equal(2, rows3[0].B);
        Assert.Throws<InvalidOperationException>(() => _excelImporter.Query(path, sheetName: "xxxx").ToList());

        using var stream = File.OpenRead(path);
        
        var rows4 = _excelImporter.Query(stream, sheetName: "Sheet3").ToList();
        Assert.Equal(5, rows4.Count);
        Assert.Equal(3, rows4[0].A);
        Assert.Equal(3, rows4[0].B);

        var rows5 = _excelImporter.Query(stream, sheetName: "Sheet2").ToList();
        Assert.Equal(12, rows5.Count);
        Assert.Equal(1, rows5[0].A);
        Assert.Equal(1, rows5[0].B);

        var rows6 = _excelImporter.Query(stream, sheetName: "Sheet1").ToList();
        Assert.Equal(12, rows6.Count);
        Assert.Equal(2, rows6[0].A);
        Assert.Equal(2, rows6[0].B);

        var rows7 = _excelImporter.Query(stream, sheetName: "Sheet1").ToList();
        Assert.Equal(12, rows7.Count);
        Assert.Equal(2, rows7[0].A);
        Assert.Equal(2, rows7[0].B);
    }

    [Fact]
    public void MultiSheetsQueryBasicTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        using var stream = File.OpenRead(path);
        
        _ = _excelImporter.Query(stream, sheetName: "Sheet1");
        _ = _excelImporter.Query(stream, sheetName: "Sheet2");
        _ = _excelImporter.Query(stream, sheetName: "Sheet3");
    }

    [Fact]
    public void MultiSheetsQueryTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheet.xlsx");
        var sheetNames1 = _excelImporter.GetSheetNames(path).ToList();
        foreach (var sheetName in sheetNames1)
        {
            var rows = _excelImporter.Query(path, sheetName: sheetName).ToList();
            Assert.NotEmpty(rows);
        }

        Assert.Equal(["Sheet1", "Sheet2", "Sheet3"], sheetNames1);

        using var stream = File.OpenRead(path);
        var sheetNames2 = _excelImporter.GetSheetNames(stream).ToList();
            
        Assert.Equal(["Sheet1", "Sheet2", "Sheet3"], sheetNames2);
        foreach (var sheetName in sheetNames2)
        {
            var rows =  _excelImporter.Query(stream, sheetName: sheetName).ToList();
            Assert.NotEmpty(rows);
        }
    }

    [Fact]
    public void ExcelSheetAttributeIsUsedWhenReadExcel()
    {
        var path = PathHelper.GetFile("xlsx/TestDynamicSheet.xlsx");
        using (var stream = File.OpenRead(path))
        {
            var users =  _excelImporter.Query<UserDto>(stream).ToList();
            Assert.Equal(2, users.Count);
            Assert.Equal("Jack", users[0].Name);

            var departments =  _excelImporter.Query<DepartmentDto>(stream).ToList();
            Assert.Equal(2, departments.Count);
            Assert.Equal("HR", departments[0].Name);
        }

        {
            var users =  _excelImporter.Query<UserDto>(path).ToList();
            Assert.Equal(2, users.Count);
            Assert.Equal("Jack", users[0].Name);

            var departments =  _excelImporter.Query<DepartmentDto>(path).ToList();
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
                new DynamicExcelSheetAttribute("usersSheet") { Name = "Users" },
                new DynamicExcelSheetAttribute("departmentSheet") { Name = "Departments" }
            ]
        };

        var path = PathHelper.GetFile("xlsx/TestDynamicSheet.xlsx");
        using (var stream = File.OpenRead(path))
        {
            // take first sheet as default
            var users =  _excelImporter.Query(stream, configuration: configuration, hasHeaderRow: true).ToList();
            Assert.Equal(2, users.Count);
            Assert.Equal("Jack", users[0].Name);

            // take second sheet by sheet name
            var departments =  _excelImporter.Query(stream, sheetName: "Departments", configuration: configuration, hasHeaderRow: true).ToList();
            Assert.Equal(2, departments.Count);
            Assert.Equal("HR", departments[0].Name);

            // take second sheet by sheet key
            departments =  _excelImporter.Query(stream, sheetName: "departmentSheet", configuration: configuration, hasHeaderRow: true).ToList();
            Assert.Equal(2, departments.Count);
            Assert.Equal("HR", departments[0].Name);
        }

        {
            // take first sheet as default
            var users =  _excelImporter.Query(path, configuration: configuration, hasHeaderRow: true).ToList();
            Assert.Equal(2, users.Count);
            Assert.Equal("Jack", users[0].Name);

            // take second sheet by sheet name
            var departments =  _excelImporter.Query(path, sheetName: "Departments", configuration: configuration, hasHeaderRow: true).ToList();
            Assert.Equal(2, departments.Count);
            Assert.Equal("HR", departments[0].Name);

            // take second sheet by sheet key
            departments =  _excelImporter.Query(path, sheetName: "departmentSheet", configuration: configuration, hasHeaderRow: true).ToList();
            Assert.Equal(2, departments.Count);
            Assert.Equal("HR", departments[0].Name);
        }
    }

    [Fact]
    public void ReadSheetVisibilityStateTest()
    {
        var path = PathHelper.GetFile("xlsx/TestMultiSheetWithHiddenSheet.xlsx");
        var sheetInfos =  _excelImporter.GetSheetInformations(path).ToList();
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

    [Fact]
    public void WriteHiddenSheetTest()
    {
        var configuration = new OpenXmlConfiguration
        {
            DynamicSheets =
            [
                new DynamicExcelSheetAttribute("usersSheet")
                {
                    Name = "Users",
                    State = SheetState.Visible
                },
                new DynamicExcelSheetAttribute("departmentSheet")
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

        var rowsWritten =  MiniExcel.Exporters.GetOpenXmlExporter().Export(path, sheets, configuration: configuration);
        Assert.Equal(2, rowsWritten.Length);
        Assert.Equal(2, rowsWritten[0]);


        var sheetInfos =  _excelImporter.GetSheetInformations(path).ToList();
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
            var rows =  _excelImporter.Query(path, sheetName: sheetName).ToList();
            Assert.NotEmpty(rows);
        }
    }
}
