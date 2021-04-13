using MiniExcelLibs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace MiniExcelTests
{
    public class MiniExcelTemplateTests
    {
        [Fact]
        public void TemplateBasiTest()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"..\..\..\..\..\samples\xlsx\TestTemplateEasyFill.xlsx";
                // 1. By POCO
                var value = new
                {
                    Name = "Jack",
                    CreateDate = new DateTime(2021, 01, 01),
                    VIP = true,
                    Points = 123
                };
                MiniExcel.SaveAsByTemplate(path, templatePath, value);

                var rows = MiniExcel.Query(path).ToList();
                Assert.Equal("Jack", rows[1].A);
                Assert.Equal("2021-01-01 00:00:00", rows[1].B);
                Assert.Equal(true, rows[1].C);
                Assert.Equal(123, rows[1].D);
                Assert.Equal("Jack has 123 points", rows[1].E);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"..\..\..\..\..\samples\xlsx\TestTemplateEasyFill.xlsx";
                // 2. By Dictionary
                var value = new Dictionary<string, object>()
                {
                    ["Name"] = "Jack",
                    ["CreateDate"] = new DateTime(2021, 01, 01),
                    ["VIP"] = true,
                    ["Points"] = 123
                };
                MiniExcel.SaveAsByTemplate(path, templatePath, value);

                var rows = MiniExcel.Query(path).ToList();
                Assert.Equal("Jack", rows[1].A);
                Assert.Equal("2021-01-01 00:00:00", rows[1].B);
                Assert.Equal(true, rows[1].C);
                Assert.Equal(123, rows[1].D);
                Assert.Equal("Jack has 123 points", rows[1].E);
            }
        }

        //[Fact]
        //public void PerformanceTest()
        //{
        //    // MiniExcel
        //    {
        //        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
        //        var templatePath = @"..\..\..\..\..\samples\xlsx\TestTemplateBasicIEmumerableFill.xlsx";
        //        var value = new
        //        {
        //            employees = Enumerable.Range(1, 1000000).Select(s => new { name = "Jack", department = "HR" })
        //        };
        //        MiniExcel.SaveAsByTemplate(path, templatePath, value);
        //    }

        //    // ClosexXml.Report
        //    {
        //        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
        //        var templatePath = @"..\..\..\..\..\samples\xlsx\TestTemplateBasicIEmumerableFill_ClosedXML_Report.xlsx";
        //        var template = new ClosedXML.Report.XLTemplate(templatePath);
        //        var value = new
        //        {
        //            employees = Enumerable.Range(1, 1000000).Select(s => new { name = "Jack", department = "HR" })
        //        };
        //        template.AddVariable(value);
        //        template.Generate();
        //        template.SaveAs(path);
        //    }
        //}

        [Fact]
        public void TestIEnumerable()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"..\..\..\..\..\samples\xlsx\TestTemplateBasicIEmumerableFill.xlsx";

                //1. By POCO
                var value = new
                {
                    employees = new[] {
                        new {name="Jack",department="HR"},
                        new {name="Lisa",department="HR"},
                        new {name="John",department="HR"},
                        new {name="Mike",department="IT"},
                        new {name="Neo",department="IT"},
                        new {name="Loan",department="IT"}
                    }
                };
                MiniExcel.SaveAsByTemplate(path, templatePath, value);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"..\..\..\..\..\samples\xlsx\TestTemplateBasicIEmumerableFill.xlsx";

                //2. By Dictionary
                var value = new Dictionary<string, object>()
                {
                    ["employees"] = new[] {
                        new {name="Jack",department="HR"},
                        new {name="Lisa",department="HR"},
                        new {name="John",department="HR"},
                        new {name="Mike",department="IT"},
                        new {name="Neo",department="IT"},
                        new {name="Loan",department="IT"}
                    }
                };
                MiniExcel.SaveAsByTemplate(path, templatePath, value);
            }
        }

        [Fact]
        public void TemplateTest()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"..\..\..\..\..\samples\xlsx\TestTemplateComplex.xlsx";

                // 1. By Class
                var value = new
                {
                    title = "FooCompany",
                    managers = new[] {
                        new {name="Jack",department="HR"},
                        new {name="Loan",department="IT"}
                    },
                    employees = new[] {
                        new {name="Wade",department="HR"},
                        new {name="Felix",department="HR"},
                        new {name="Eric",department="IT"},
                        new {name="Keaton",department="IT"}
                    }
                };
                MiniExcel.SaveAsByTemplate(path, templatePath, value);

                var rows = MiniExcel.Query(path).ToList();
                Assert.Equal("FooCompany", rows[0].A);
                Assert.Equal("Jack", rows[2].B);
                Assert.Equal("HR", rows[2].C);
                Assert.Equal("Loan", rows[3].B);
                Assert.Equal("IT", rows[3].C);

                Assert.Equal("Wade", rows[5].B);
                Assert.Equal("HR", rows[5].C);
                Assert.Equal("Felix", rows[6].B);
                Assert.Equal("HR", rows[6].C);

                Assert.Equal("Eric", rows[7].B);
                Assert.Equal("IT", rows[7].C);
                Assert.Equal("Keaton", rows[8].B);
                Assert.Equal("IT", rows[8].C);
            }


            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"..\..\..\..\..\samples\xlsx\TestTemplateComplex.xlsx";

                // 2. By Dictionary
                var value = new Dictionary<string, object>()
                {
                    ["title"] = "FooCompany",
                    ["managers"] = new[] {
                    new {name="Jack",department="HR"},
                    new {name="Loan",department="IT"}
                },
                    ["employees"] = new[] {
                    new {name="Wade",department="HR"},
                    new {name="Felix",department="HR"},
                    new {name="Eric",department="IT"},
                    new {name="Keaton",department="IT"}
                }
                };
                MiniExcel.SaveAsByTemplate(path, templatePath, value);

                var rows = MiniExcel.Query(path).ToList();
                Assert.Equal("FooCompany", rows[0].A);
                Assert.Equal("Jack", rows[2].B);
                Assert.Equal("HR", rows[2].C);
                Assert.Equal("Loan", rows[3].B);
                Assert.Equal("IT", rows[3].C);

                Assert.Equal("Wade", rows[5].B);
                Assert.Equal("HR", rows[5].C);
                Assert.Equal("Felix", rows[6].B);
                Assert.Equal("HR", rows[6].C);

                Assert.Equal("Eric", rows[7].B);
                Assert.Equal("IT", rows[7].C);
                Assert.Equal("Keaton", rows[8].B);
                Assert.Equal("IT", rows[8].C);
            }

        }
    }
}
