using CsvHelper;
using CsvHelper.Configuration;
using MiniExcelLibs.Tests.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace MiniExcelLibs.Tests
{
    public class MiniExcelCsvTests
    {
        [Fact]
        public void gb2312_Encoding_Read_Test()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var path = PathHelper.GetSamplePath("csv/gb2312_Encoding_Read_Test.csv");
            var config = new MiniExcelLibs.Csv.CsvConfiguration()
            {
                StreamReaderFunc = (stream) => new StreamReader(stream,encoding: Encoding.GetEncoding("gb2312"))
            };
            var rows = MiniExcel.Query(path, true,excelType:ExcelType.CSV,configuration: config).ToList();
            Assert.Equal("世界你好", rows[0].栏位1);
        }

        [Fact]
        public void SeperatorTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.csv");
            var values = new List<Dictionary<string, object>>()
                {
                    new Dictionary<string,object>{{ "a", @"""<>+-*//}{\\n" }, { "b", 1234567890 },{ "c", true },{ "d", new DateTime(2021, 1, 1) } },
                    new Dictionary<string,object>{{ "a", @"<test>Hello World</test>" }, { "b", -1234567890 },{ "c", false },{ "d", new DateTime(2021, 1, 2) } },
                };
            MiniExcel.SaveAs(path, values,configuration: new MiniExcelLibs.Csv.CsvConfiguration() {Seperator=';'});
            var expected = @"a;b;c;d
""""""<>+-*//}{\\n"";1234567890;True;""2021-01-01 00:00:00""
""<test>Hello World</test>"";-1234567890;False;""2021-01-02 00:00:00""
";
            Assert.Equal(expected, File.ReadAllText(path));
        }

        [Fact]
        public void SaveAsByDictionary()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.csv");
                var table = new List<Dictionary<string, object>>();
                MiniExcel.SaveAs(path, table);
                Assert.Equal("\r\n", File.ReadAllText(path));
                File.Delete(path);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.csv");
                var table = new Dictionary<string, object>(); //TODO
                MiniExcel.SaveAs(path, table);
                //Assert.Throws<NotImplementedException>(() => MiniExcel.SaveAs(path, table));
                Assert.Equal("\r\n", File.ReadAllText(path));
                File.Delete(path);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.csv");
                var values = new List<Dictionary<string, object>>()
                {
                    new Dictionary<string,object>{{ "a", @"""<>+-*//}{\\n" }, { "b", 1234567890 },{ "c", true },{ "d", new DateTime(2021, 1, 1) } },
                    new Dictionary<string,object>{{ "a", @"<test>Hello World</test>" }, { "b", -1234567890 },{ "c", false },{ "d", new DateTime(2021, 1, 2) } },
                };
                MiniExcel.SaveAs(path, values);

                using (var reader = new StreamReader(path))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var records = csv.GetRecords<dynamic>().ToList();
                    Assert.Equal(@"""<>+-*//}{\\n", records[0].a);
                    Assert.Equal(@"1234567890", records[0].b);
                    Assert.Equal(@"True", records[0].c);
                    Assert.Equal(@"2021-01-01 00:00:00", records[0].d);

                    Assert.Equal(@"<test>Hello World</test>", records[1].a);
                    Assert.Equal(@"-1234567890", records[1].b);
                    Assert.Equal(@"False", records[1].c);
                    Assert.Equal(@"2021-01-02 00:00:00", records[1].d);
                }

                File.Delete(path);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.csv");
                var values = new List<Dictionary<int, object>>()
                {
                    new Dictionary<int,object>{{ 1, @"""<>+-*//}{\\n" }, { 2, 1234567890 },{ 3, true },{ 4, new DateTime(2021, 1, 1) } },
                    new Dictionary<int,object>{{ 1, @"<test>Hello World</test>" }, { 2, -1234567890 },{ 3, false },{4, new DateTime(2021, 1, 2) } },
                };
                MiniExcel.SaveAs(path, values);

                using (var reader = new StreamReader(path))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var records = csv.GetRecords<dynamic>().ToList();
                    {
                        var row = records[0] as IDictionary<string, object>;
                        Assert.Equal(@"""<>+-*//}{\\n", row["1"]);
                        Assert.Equal(@"1234567890", row["2"]);
                        Assert.Equal(@"True", row["3"]);
                        Assert.Equal(@"2021-01-01 00:00:00", row["4"]);
                    }
                    {
                        var row = records[1] as IDictionary<string, object>;
                        Assert.Equal(@"<test>Hello World</test>", row["1"]);
                        Assert.Equal(@"-1234567890", row["2"]);
                        Assert.Equal(@"False", row["3"]);
                        Assert.Equal(@"2021-01-02 00:00:00", row["4"]);
                    }
                }

                File.Delete(path);
            }
        }

        [Fact]
        public void SaveAsByDataTableTest()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.csv");
                var table = new DataTable();
                MiniExcel.SaveAs(path, table);

                var text = File.ReadAllText(path);
                Assert.Equal("\r\n", text);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.csv");

                var table = new DataTable();
                {
                    table.Columns.Add("a", typeof(string));
                    table.Columns.Add("b", typeof(decimal));
                    table.Columns.Add("c", typeof(bool));
                    table.Columns.Add("d", typeof(DateTime));
                    table.Rows.Add(@"""<>+-*//}{\\n", 1234567890, true, new DateTime(2021, 1, 1));
                    table.Rows.Add(@"<test>Hello World</test>", -1234567890, false, new DateTime(2021, 1, 2));
                }

                MiniExcel.SaveAs(path, table);

                using (var reader = new StreamReader(path))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var records = csv.GetRecords<dynamic>().ToList();
                    Assert.Equal(@"""<>+-*//}{\\n", records[0].a);
                    Assert.Equal(@"1234567890", records[0].b);
                    Assert.Equal(@"True", records[0].c);
                    Assert.Equal(@"2021-01-01 00:00:00", records[0].d);

                    Assert.Equal(@"<test>Hello World</test>", records[1].a);
                    Assert.Equal(@"-1234567890", records[1].b);
                    Assert.Equal(@"False", records[1].c);
                    Assert.Equal(@"2021-01-02 00:00:00", records[1].d);
                }

                File.Delete(path);
            }

        }


        public class Test
        {
            public string c1 { get; set; }
		  public string c2 { get; set; }
	   }

        [Fact]
        public void CsvExcelTypeTest()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.csv");
                var input = new[] { new { A = "Test1", B = "Test2" } };
                MiniExcel.SaveAs(path, input);

                var texts = File.ReadAllLines(path);
                Assert.Equal("A,B", texts[0]);
                Assert.Equal("Test1,Test2", texts[1]);

                {
                    var rows = MiniExcel.Query(path).ToList();
                    Assert.Equal("A", rows[0].A);
                    Assert.Equal("B", rows[0].B);
                    Assert.Equal("Test1", rows[1].A);
                    Assert.Equal("Test2", rows[1].B);
                }

                using (var reader = new StreamReader(path))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var rows = csv.GetRecords<dynamic>().ToList();
                    Assert.Equal("Test1", rows[0].A);
                    Assert.Equal("Test2", rows[0].B);
                }

                File.Delete(path);
            }
        }

        [Fact()]
	   public void Create2x2_Test()
        {
		  var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.csv");
		  MiniExcel.SaveAs(path, new[] {
			 new { c1 = "A1" ,c2 = "B1"},
			 new { c1 = "A2" ,c2 = "B2"},
		  });

            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query(useHeaderRow: true,excelType:ExcelType.CSV).ToList();
                Assert.Equal("A1", rows[0].c1);
                Assert.Equal("B1", rows[0].c2);
                Assert.Equal("A2", rows[1].c1);
                Assert.Equal("B2", rows[1].c2);
            }

		  {
			 var rows = MiniExcel.Query(path,useHeaderRow: true, excelType: ExcelType.CSV).ToList();
			 Assert.Equal("A1", rows[0].c1);
			 Assert.Equal("B1", rows[0].c2);
			 Assert.Equal("A2", rows[1].c1);
			 Assert.Equal("B2", rows[1].c2);
		  }

            File.Delete(path);
        }

        [Fact()]
        public void CsvTypeMappingTest()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.csv");
            MiniExcel.SaveAs(path, new[] {
                new { c1 = "A1" ,c2 = "B1"},
                new { c1 = "A2" ,c2 = "B2"},
            });

            using (var stream = File.OpenRead(path))
            {
                var rows = stream.Query<Test>(excelType: ExcelType.CSV).ToList();
                Assert.Equal("A1", rows[0].c1);
                Assert.Equal("B1", rows[0].c2);
                Assert.Equal("A2", rows[1].c1);
                Assert.Equal("B2", rows[1].c2);
            }

		  {
			 var rows = MiniExcel.Query<Test>(path, excelType: ExcelType.CSV).ToList();
			 Assert.Equal("A1", rows[0].c1);
			 Assert.Equal("B1", rows[0].c2);
			 Assert.Equal("A2", rows[1].c1);
			 Assert.Equal("B2", rows[1].c2);
		  }

            File.Delete(path);
        }

        [Fact()]
	   public void Delimiters_Test()
        {
		  //TODO:Datetime have default format like yyyy-MM-dd HH:mm:ss ?
		  {
			 Assert.Equal(Generate("\"\"\""), MiniExcelGenerateCsv("\"\"\""));
			 Assert.Equal(Generate(","), MiniExcelGenerateCsv(","));
			 Assert.Equal(Generate(" "), MiniExcelGenerateCsv(" "));
			 Assert.Equal(Generate(";"), MiniExcelGenerateCsv(";"));
			 Assert.Equal(Generate("\t"), MiniExcelGenerateCsv("\t"));
		  }
	   }

	   string Generate(string value)
	   {
		  var records = Enumerable.Range(1, 1).Select((s, idx) => new { v1 = value, v2 = value });
		  var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.csv");
		  using (var writer = new StreamWriter(path))
		  using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
		  {
			 csv.WriteRecords(records);
		  }

		  var content = File.ReadAllText(path);
		  File.Delete(path);
		  return content;
	   }

	   string MiniExcelGenerateCsv(string value)
	   {
		  var records = Enumerable.Range(1, 1).Select((s, idx) => new { v1 = value, v2 = value });
		  var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.csv");

            using (var stream = File.Create(path))
            {
                stream.SaveAs(records,excelType:ExcelType.CSV);
            }

		  var content = File.ReadAllText(path);
		  File.Delete(path);
		  return content;
	   }
    }
}