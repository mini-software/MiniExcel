using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace MiniExcelLibs.Tests
{
    public partial class MiniExcelOpenXmlMultipleSheetTests
    {
        [Fact]
        public void SpecifySheetNameQueryTest()
        {
            var path = @"../../../../../samples/xlsx/TestMultiSheet.xlsx";
            {
                var rows = MiniExcel.Query(path, sheetName: "Sheet3").ToList();
                Assert.Equal(5, rows.Count);
                Assert.Equal(3, rows[0].A);
                Assert.Equal(3, rows[0].B);
            }
            {
                var rows = MiniExcel.Query(path, sheetName: "Sheet2").ToList();
                Assert.Equal(12, rows.Count);
                Assert.Equal(1, rows[0].A);
                Assert.Equal(1, rows[0].B);
            }
            {
                var rows = MiniExcel.Query(path, sheetName: "Sheet1").ToList();
                Assert.Equal(12, rows.Count);
                Assert.Equal(2, rows[0].A);
                Assert.Equal(2, rows[0].B);
            }
            {
                Assert.Throws<InvalidOperationException>(() => MiniExcel.Query(path, sheetName: "xxxx").ToList());
            }

            using (var stream = File.OpenRead(path))
            {
                {
                    var rows = stream.Query(sheetName: "Sheet3").ToList();
                    Assert.Equal(5, rows.Count);
                    Assert.Equal(3, rows[0].A);
                    Assert.Equal(3, rows[0].B);
                }
                {
                    var rows = stream.Query(sheetName: "Sheet2").ToList();
                    Assert.Equal(12, rows.Count);
                    Assert.Equal(1, rows[0].A);
                    Assert.Equal(1, rows[0].B);
                }
                {
                    var rows = stream.Query(sheetName: "Sheet1").ToList();
                    Assert.Equal(12, rows.Count);
                    Assert.Equal(2, rows[0].A);
                    Assert.Equal(2, rows[0].B);
                }
                {
                    var rows = stream.Query(sheetName: "Sheet1").ToList();
                    Assert.Equal(12, rows.Count);
                    Assert.Equal(2, rows[0].A);
                    Assert.Equal(2, rows[0].B);
                }
            }
        }

        [Fact]
        public void MultiSheetsQueryBasicTest()
        {
            var path = @"../../../../../samples/xlsx/TestMultiSheet.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var sheet1 = stream.Query(sheetName: "Sheet1");
                var sheet2 = stream.Query(sheetName: "Sheet2");
                var sheet3 = stream.Query(sheetName: "Sheet3");
            }
        }

        [Fact]
        public void MultiSheetsQueryTest()
        {
            var path = @"../../../../../samples/xlsx/TestMultiSheet.xlsx";
            {
                var sheetNames = MiniExcel.GetSheetNames(path).ToList();
                foreach (var sheetName in sheetNames)
                {
                    var rows = MiniExcel.Query(path, sheetName: sheetName);
                }

                Assert.Equal(new[] { "Sheet2", "Sheet1", "Sheet3" }, sheetNames);
            }

            {
                using (var stream = File.OpenRead(path))
                {
                    var sheetNames = stream.GetSheetNames().ToList();
                    Assert.Equal(new[] { "Sheet2", "Sheet1", "Sheet3" }, sheetNames);
                    foreach (var sheetName in sheetNames)
                    {
                        var rows = stream.Query(sheetName: sheetName);
                    }
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
            var path = @"../../../../../samples/xlsx/TestDynamicSheet.xlsx";
            using (var stream = File.OpenRead(path))
            {
                var users = stream.Query<UserDto>().ToList();
                Assert.Equal(2, users.Count());
                Assert.Equal("Jack", users[0].Name);

                var departments = stream.Query<DepartmentDto>().ToList();
                Assert.Equal(2, departments.Count());
                Assert.Equal("HR", departments[0].Name);
            }

            {
                var users = MiniExcel.Query<UserDto>(path).ToList();
                Assert.Equal(2, users.Count());
                Assert.Equal("Jack", users[0].Name);

                var departments = MiniExcel.Query<DepartmentDto>(path).ToList();
                Assert.Equal(2, departments.Count());
                Assert.Equal("HR", departments[0].Name);
            }
        }

        [Fact]
        public void DynamicSheetConfigurationIsUsedWhenReadExcel()
        {
            var configuration = new OpenXmlConfiguration
            {
                DynamicSheets = new[]
                {
                    new DynamicExcelSheet("usersSheet") { Name = "Users" },
                    new DynamicExcelSheet("departmentSheet") { Name = "Departments" }
                }
            };

            var path = @"../../../../../samples/xlsx/TestDynamicSheet.xlsx";
            using (var stream = File.OpenRead(path))
            {
                // take first sheet as default
                var users = stream.Query(configuration: configuration, useHeaderRow: true).ToList();
                Assert.Equal(2, users.Count());
                Assert.Equal("Jack", users[0].Name);

                // take second sheet by sheet name
                var departments = stream.Query(sheetName: "Departments", configuration: configuration, useHeaderRow: true).ToList();
                Assert.Equal(2, departments.Count());
                Assert.Equal("HR", departments[0].Name);

                // take second sheet by sheet key
                departments = stream.Query(sheetName: "departmentSheet", configuration: configuration, useHeaderRow: true).ToList();
                Assert.Equal(2, departments.Count());
                Assert.Equal("HR", departments[0].Name);
            }

            {
                // take first sheet as default
                var users = MiniExcel.Query(path, configuration: configuration, useHeaderRow: true).ToList();
                Assert.Equal(2, users.Count());
                Assert.Equal("Jack", users[0].Name);

                // take second sheet by sheet name
                var departments = MiniExcel.Query(path, sheetName: "Departments", configuration: configuration, useHeaderRow: true).ToList();
                Assert.Equal(2, departments.Count());
                Assert.Equal("HR", departments[0].Name);

                // take second sheet by sheet key
                departments = MiniExcel.Query(path, sheetName: "departmentSheet", configuration: configuration, useHeaderRow: true).ToList();
                Assert.Equal(2, departments.Count());
                Assert.Equal("HR", departments[0].Name);
            }
        }

        [Fact]
        public void ReadSheetVisibilityStateTest()
        {
            var path = @"../../../../../samples/xlsx/TestMultiSheetWithHiddenSheet.xlsx";
            {
                var sheetInfos = MiniExcel.GetSheetInformations(path).ToList();
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
                DynamicSheets = new[]
                {
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
                }
            };

            var users = new[] { new { Name = "Jack", Age = 25 }, new { Name = "Mike", Age = 44 } };
            var department = new[] { new { ID = "01", Name = "HR" }, new { ID = "02", Name = "IT" } };
            var sheets = new Dictionary<string, object>
            {
                ["usersSheet"] = users,
                ["departmentSheet"] = department
            };

            string path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
            MiniExcel.SaveAs(path, sheets, configuration: configuration);

            var sheetInfos = MiniExcel.GetSheetInformations(path).ToList();
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
                var rows = MiniExcel.Query(path, sheetName: sheetName);
            }
        }
    }
}