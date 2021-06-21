using Dapper;
using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Tests.Utils;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;
using static MiniExcelLibs.Tests.MiniExcelOpenXmlTests;

namespace MiniExcelLibs.Tests
{
    public partial class MiniExcelIssueAsyncTests
    {
        private readonly ITestOutputHelper output;
        public MiniExcelIssueAsyncTests(ITestOutputHelper output)
        {
            this.output = output;
        }

        /// <summary>
        /// [SaveAsByTemplate support DateTime custom format · Issue #255 · shps951023/MiniExcel]
        /// (https://github.com/shps951023/MiniExcel/issues/255)
        /// </summary>
        [Fact]
        public async Task Issue255()
        {
            //tempalte
            {
                var templatePath = PathHelper.GetSamplePath("xlsx/TestsIssue255_Template.xlsx");
                var path = PathHelper.GetTempPath();
                var value = new
                {
                    Issue255DTO = new Issue255DTO[] {
                        new Issue255DTO { Time = new DateTime(2021, 01, 01) }
                    }
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);
                var q = await MiniExcel.QueryAsync(path);
                var rows = q.ToList();
                Assert.Equal("2021", rows[1].A.ToString());
            }
            //saveas
            {
                var path = PathHelper.GetTempPath();
                var value = new Issue255DTO[] {
                    new Issue255DTO { Time = new DateTime(2021, 01, 01) }
                };
                await MiniExcel.SaveAsAsync(path, value);
                var q = await MiniExcel.QueryAsync(path);
                var rows = q.ToList();
                Assert.Equal("2021", rows[1].A.ToString());
            }
        }

        public class Issue255DTO
        {
            [ExcelFormat("yyyy")]
            public DateTime Time { get; set; }
        }

        /// <summary>
        /// [Dynamic QueryAsync custom format not using mapping format · Issue #256]
        /// (https://github.com/shps951023/MiniExcel/issues/256)
        /// </summary>
        [Fact]
        public async Task Issue256()
        {
            var path = PathHelper.GetSamplePath("xlsx/TestIssue256.xlsx");
            var q = await MiniExcel.QueryAsync(path, false);
            var rows = q.ToList();
            Assert.Equal(new DateTime(2003, 4, 16), rows[1].A);
            Assert.Equal(new DateTime(2004, 4, 16), rows[1].B);
        }


        /// <summary>
        /// Csv SaveAs by datareader with encoding default show messy code #253
        /// </summary>
        [Fact]
        public async Task Issue253()
        {
            {
                var value = new[] { new { col1 = "世界你好" } };
                var path = PathHelper.GetTempPath(extension: "csv");
                await MiniExcel.SaveAsAsync(path, value);
                var expected = @"col1
世界你好
";
                Assert.Equal(expected, File.ReadAllText(path));
            }

            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                var value = new[] { new { col1 = "世界你好" } };
                var path = PathHelper.GetTempPath(extension: "csv");
                var config = new MiniExcelLibs.Csv.CsvConfiguration()
                {
                    StreamWriterFunc = (stream) => new StreamWriter(stream, Encoding.GetEncoding("gb2312"))
                };
                await MiniExcel.SaveAsAsync(path, value, excelType: ExcelType.CSV, configuration: config);
                var expected = @"col1
�������
";
                Assert.Equal(expected, File.ReadAllText(path));
            }

            using (var cn = Db.GetConnection())
            {
                var value = cn.ExecuteReader(@"select '世界你好' col1");
                var path = PathHelper.GetTempPath(extension: "csv");
                await MiniExcel.SaveAsAsync(path, value);
                var expected = @"col1
世界你好
";
                Assert.Equal(expected, File.ReadAllText(path));
            }
        }

        /// <summary>
        /// [CSV SaveAs support datareader · Issue #251 · shps951023/MiniExcel](https://github.com/shps951023/MiniExcel/issues/251)
        /// </summary>
        [Fact]
        public async Task Issue251()
        {
            using (var cn = Db.GetConnection())
            {
                var reader = await cn.ExecuteReaderAsync(@"select '""<>+-*//}{\\n' a,1234567890 b union all select '<test>Hello World</test>',-1234567890");
                var path = PathHelper.GetTempPath(extension: "csv");
                await MiniExcel.SaveAsAsync(path, reader);
                var expected = @"a,b
""""""<>+-*//}{\\n"",1234567890
""<test>Hello World</test>"",-1234567890
";
                Assert.Equal(expected, File.ReadAllText(path));
            }
        }

        /// <summary>
        /// No error exception throw when reading xls file #242
        /// </summary>
        [Fact]
        public async Task Issue242()
        {
            var path = PathHelper.GetSamplePath("xls/TestIssue242.xls");

            await Assert.ThrowsAsync<NotSupportedException>(async () => {
                var q = await MiniExcel.QueryAsync(path);
                q.ToList();
            });

            using (var stream = File.OpenRead(path))
            {
                await Assert.ThrowsAsync<NotSupportedException>(async () => {
                    var q = await  stream.QueryAsync();
                    q.ToList();
                });
            }
        }

        /// <summary>
        /// Csv type mapping QueryAsync error "cannot be converted to xxx type" #243
        /// </summary>
        [Fact]
        public async Task Issue243()
        {
            var path = PathHelper.GetTempPath("csv");
            var value = new[] {
                  new { name ="Jack",Age=25,InDate=new DateTime(2021,01,03)},
                  new { name ="Henry",Age=36,InDate=new DateTime(2020,05,03)},
             };
            await MiniExcel.SaveAsAsync(path, value);

            var q = await MiniExcel.QueryAsync<Issue243Dto>(path);
            var rows = q.ToList();
            Assert.Equal("Jack", rows[0].name);
            Assert.Equal(25, rows[0].Age);
            Assert.Equal(new DateTime(2021, 01, 03), rows[0].InDate);

            Assert.Equal("Henry", rows[1].name);
            Assert.Equal(36, rows[1].Age);
            Assert.Equal(new DateTime(2020, 05, 03), rows[1].InDate);
        }

        public class Issue243Dto
        {
            public string name { get; set; }
            public int Age { get; set; }
            public DateTime InDate { get; set; }
        }

        /// <summary>
        /// Support Custom Datetime format #241
        /// </summary>
        [Fact]
        public async Task Issue241()
        {

            var value = new Issue241Dto[] {
                new Issue241Dto{ Name="Jack",InDate=new DateTime(2021,01,04)},
                new Issue241Dto{ Name="Henry",InDate=new DateTime(2020,04,05)},
            };

            // csv
            {
                var path = PathHelper.GetTempPath("csv");
                MiniExcel.SaveAs(path, value);

                {
                    var q = await MiniExcel.QueryAsync(path, true);
                    var rows = q.ToList();
                    Assert.Equal(rows[0].InDate, "01 04, 2021");
                    Assert.Equal(rows[1].InDate, "04 05, 2020");
                }

                {
                    var q = await MiniExcel.QueryAsync<Issue241Dto>(path);
                    var rows = q.ToList(); 
                    Assert.Equal(rows[0].InDate, new DateTime(2021, 01, 04));
                    Assert.Equal(rows[1].InDate, new DateTime(2020, 04, 05));
                }
            }

            // xlsx
            {
                var path = PathHelper.GetTempPath();
                await MiniExcel.SaveAsAsync(path, value);

                {
                    var q = await MiniExcel.QueryAsync(path, true);
                    var rows = q.ToList();
                    Assert.Equal(rows[0].InDate, "01 04, 2021");
                    Assert.Equal(rows[1].InDate, "04 05, 2020");
                }

                {
                    var q = await MiniExcel.QueryAsync<Issue241Dto>(path);
                    var rows = q.ToList();
                    Assert.Equal(rows[0].InDate, new DateTime(2021, 01, 04));
                    Assert.Equal(rows[1].InDate, new DateTime(2020, 04, 05));
                }
            }
        }

        public class Issue241Dto
        {
            public string Name { get; set; }

            [ExcelFormat("MM dd, yyyy")]
            public DateTime InDate { get; set; }
        }

        /// <summary>
        /// SaveAs Default Template #132
        /// </summary>
        [Fact]
        public async Task Issue132()
        {
            {
                var path = PathHelper.GetTempPath();
                var value = new[] {
                    new { name ="Jack",Age=25,InDate=new DateTime(2021,01,03)},
                    new { name ="Henry",Age=36,InDate=new DateTime(2020,05,03)},
                };

                await MiniExcel.SaveAsAsync(path, value);
            }

            {
                var path = PathHelper.GetTempPath();
                var value = new[] {
                    new { name ="Jack",Age=25,InDate=new DateTime(2021,01,03)},
                    new { name ="Henry",Age=36,InDate=new DateTime(2020,05,03)},
                };
                var config = new OpenXmlConfiguration()
                {
                    TableStyles = TableStyles.None
                };
                await MiniExcel.SaveAsAsync(path, value, configuration: config);
            }

            {
                var path = PathHelper.GetTempPath();
                var value = JsonConvert.DeserializeObject<DataTable>(
                    JsonConvert.SerializeObject(new[] {
                        new { name ="Jack",Age=25,InDate=new DateTime(2021,01,03)},
                        new { name ="Henry",Age=36,InDate=new DateTime(2020,05,03)},
                    })
                );
                await MiniExcel.SaveAsAsync(path, value);
            }
        }

        /// <summary>
        /// Support SaveAs by DataSet #235
        /// </summary>
        [Fact]
        public async Task Issue235()
        {
            var path = PathHelper.GetTempPath();

            DataSet sheets = new DataSet();
            var users = JsonConvert.DeserializeObject<DataTable>(JsonConvert.SerializeObject(new[] { new { Name = "Jack", Age = 25 }, new { Name = "Mike", Age = 44 } }));
            users.TableName = "users";
            var department = JsonConvert.DeserializeObject<DataTable>(JsonConvert.SerializeObject(new[] { new { ID = "01", Name = "HR" }, new { ID = "02", Name = "IT" } })); ;
            department.TableName = "department";
            sheets.Tables.Add(users);
            sheets.Tables.Add(department);

            await MiniExcel.SaveAsAsync(path, sheets);


            var sheetNames = MiniExcel.GetSheetNames(path);
            Assert.Equal("users", sheetNames[0]);
            Assert.Equal("department", sheetNames[1]);

            {
                var q = await MiniExcel.QueryAsync(path, true, sheetName: "users");
                var rows = q.ToList();
                Assert.Equal("Jack", rows[0].Name);
                Assert.Equal(25, rows[0].Age);
                Assert.Equal("Mike", rows[1].Name);
                Assert.Equal(44, rows[1].Age);
            }
            {
                var q = await MiniExcel.QueryAsync(path, true, sheetName: "department");
                var rows = q.ToList();
                Assert.Equal("01", rows[0].ID);
                Assert.Equal("HR", rows[0].Name);
                Assert.Equal("02", rows[1].ID);
                Assert.Equal("IT", rows[1].Name);
            }
        }

        /// <summary>
        /// QueryAsDataTable A2=5.5 , A3=0.55/1.1 will case double type check error #233
        /// </summary>
        [Fact]
        public async Task Issue233()
        {
            var path = PathHelper.GetSamplePath("xlsx/TestIssue233.xlsx");
            var dt = await MiniExcel.QueryAsDataTableAsync(path);
            var rows = dt.Rows;

            Assert.Equal(0.55, rows[0]["Size"]);
            Assert.Equal("0.55/1.1", rows[1]["Size"]);
        }

        /// <summary>
        /// Csv QueryAsync split comma not correct #237
        /// https://github.com/shps951023/MiniExcel/issues/237
        /// </summary>
        [Fact]
        public async Task Issue237()
        {
            var value = new[]
            {
                new{ id="\"\"1,2,3\"\""},
                new{ id="1,2,3"},
            };
            var path = PathHelper.GetTempPath("csv");
            await MiniExcel.SaveAsAsync(path, value);

            var q = await MiniExcel.QueryAsync(path, true);
            var rows = q.ToList();
            Assert.Equal("\"\"1,2,3\"\"", rows[0].id);
            Assert.Equal("1,2,3", rows[1].id);
        }

        /// <summary>
        /// SaveAs support multiple sheets #234
        /// </summary>
        [Fact]
        public async Task Issue234()
        {
            var path = PathHelper.GetTempPath();

            var users = new[] { new { Name = "Jack", Age = 25 }, new { Name = "Mike", Age = 44 } };
            var department = new[] { new { ID = "01", Name = "HR" }, new { ID = "02", Name = "IT" } };
            var sheets = new Dictionary<string, object>
            {
                ["users"] = users,
                ["department"] = department
            };
            await MiniExcel.SaveAsAsync(path, sheets);

            var sheetNames = MiniExcel.GetSheetNames(path);
            Assert.Equal("users", sheetNames[0]);
            Assert.Equal("department", sheetNames[1]);

            {
                var q = await MiniExcel.QueryAsync(path, true, sheetName: "users");
                var rows = q.ToList();
                Assert.Equal("Jack", rows[0].Name);
                Assert.Equal(25, rows[0].Age);
                Assert.Equal("Mike", rows[1].Name);
                Assert.Equal(44, rows[1].Age);
            }
            {
                var q = await MiniExcel.QueryAsync(path, true, sheetName: "department");
                var rows = q.ToList();
                Assert.Equal("01", rows[0].ID);
                Assert.Equal("HR", rows[0].Name);
                Assert.Equal("02", rows[1].ID);
                Assert.Equal("IT", rows[1].Name);
            }
        }

        /// <summary>
        /// SaveAs By Reader Closed error : 'Error! Invalid attempt to call FieldCount when reader is closed' #230
        /// https://github.com/shps951023/MiniExcel/issues/230
        /// </summary>
        [Fact]
        public async Task Issue230()
        {
            var conn = Db.GetConnection("Data Source=:memory:");
            conn.Open();
            var cmd = conn.CreateCommand();
            cmd.CommandText = "select 1 id union all select 2";
            using (var reader = await cmd.ExecuteReaderAsync(CommandBehavior.CloseConnection))
            {
                while (reader.Read())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        var result = $"{reader.GetName(i)} , {reader.GetValue(i)}";
                        output.WriteLine(result);
                    }
                }
            }

            conn = Db.GetConnection("Data Source=:memory:");
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandText = "select 1 id union all select 2";
            using (var reader = await cmd.ExecuteReaderAsync(CommandBehavior.CloseConnection))
            {
                while (reader.Read())
                {
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        var result = $"{reader.GetName(i)} , {reader.GetValue(i)}";
                        output.WriteLine(result);
                    }
                }
            }

            conn = Db.GetConnection("Data Source=:memory:");
            conn.Open();
            cmd = conn.CreateCommand();
            cmd.CommandText = "select 1 id union all select 2";
            using (var reader = await cmd.ExecuteReaderAsync(CommandBehavior.CloseConnection))
            {
                var path = PathHelper.GetTempPath();
                await MiniExcel.SaveAsAsync(path, reader, printHeader: true);
                var q = await MiniExcel.QueryAsync(path, true);
                var rows = q.ToList();
                Assert.Equal(1, rows[0].id);
                Assert.Equal(2, rows[1].id);
            }
        }

        /// <summary>
        /// v0.14.3 QueryAsDataTable error "Cannot set Column to be null" #229
        /// https://github.com/shps951023/MiniExcel/issues/229
        /// </summary>
        [Fact]
        public async Task Issue229()
        {
            var path = PathHelper.GetSamplePath("xlsx/TestIssue229.xlsx");
            var dt = await MiniExcel.QueryAsDataTableAsync(path);
            foreach (DataColumn column in dt.Columns)
            {
                var v = dt.Rows[3][column];
                var type = v?.GetType();
                Assert.Equal(DBNull.Value, v);
            }
        }

        /// <summary>
        /// [QueryAsync Merge cells data · Issue #122 · shps951023/MiniExcel]
        /// (https://github.com/shps951023/MiniExcel/issues/122)
        /// </summary>
        [Fact]
        public async Task Issue122()
        {
            var config = new OpenXmlConfiguration()
            {
                FillMergedCells = true
            };
            {
                var path = PathHelper.GetSamplePath("xlsx/TestIssue122.xlsx");
                {
                    var q = await MiniExcel.QueryAsync(path, useHeaderRow: true, configuration: config);
                    var rows = q.ToList();
                    Assert.Equal("HR", rows[0].Department);
                    Assert.Equal("HR", rows[1].Department);
                    Assert.Equal("HR", rows[2].Department);
                    Assert.Equal("IT", rows[3].Department);
                    Assert.Equal("IT", rows[4].Department);
                    Assert.Equal("IT", rows[5].Department);
                }
            }

            {
                var path = PathHelper.GetSamplePath("xlsx/TestIssue122_2.xlsx");
                {
                    var q = await MiniExcel.QueryAsync(path, useHeaderRow: true, configuration: config);
                    var rows = q.ToList();
                    Assert.Equal("V1", rows[2].Test1);
                    Assert.Equal("V2", rows[5].Test2);
                    Assert.Equal("V3", rows[1].Test3);
                    Assert.Equal("V4", rows[2].Test4);
                    Assert.Equal("V5", rows[3].Test5);
                    Assert.Equal("V6", rows[5].Test5);
                }
            }
        }

        /// <summary>
        /// [Support Xlsm AutoCheck · Issue #227 · shps951023/MiniExcel]
        /// (https://github.com/shps951023/MiniExcel/issues/227)
        /// </summary>
        [Fact]
        public async Task Issue227()
        {
            {
                var path = PathHelper.GetTempPath("xlsm");
                Assert.Throws<NotSupportedException>(() => MiniExcel.SaveAs(path, new[] { new { V = "A1" }, new { V = "A2" } }));
            }

            {
                var path = PathHelper.GetSamplePath("xlsx/TestIssue227.xlsm");
                {
                    var q = await MiniExcel.QueryAsync<UserAccount>(path);
                    var rows = q.ToList();
                    Assert.Equal(100, rows.Count());

                    Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows[0].ID);
                    Assert.Equal("Wade", rows[0].Name);
                    Assert.Equal(DateTime.ParseExact("27/09/2020", "dd/MM/yyyy", CultureInfo.InvariantCulture), rows[0].BoD);
                    Assert.Equal(36, rows[0].Age);
                    Assert.False(rows[0].VIP);
                    Assert.Equal(decimal.Parse("5019.12"), rows[0].Points);
                    Assert.Equal(1, rows[0].IgnoredProperty);
                }
                {
                    using (var stream = File.OpenRead(path))
                    {
                        var q = await stream.QueryAsync<UserAccount>();
                        var rows = q.ToList();
                        Assert.Equal(100, rows.Count());

                        Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows[0].ID);
                        Assert.Equal("Wade", rows[0].Name);
                        Assert.Equal(DateTime.ParseExact("27/09/2020", "dd/MM/yyyy", CultureInfo.InvariantCulture), rows[0].BoD);
                        Assert.Equal(36, rows[0].Age);
                        Assert.False(rows[0].VIP);
                        Assert.Equal(decimal.Parse("5019.12"), rows[0].Points);
                        Assert.Equal(1, rows[0].IgnoredProperty);
                    }
                }
            }


        }

        /// <summary>
        /// https://github.com/shps951023/MiniExcel/issues/226
        /// Fix SaveAsByTemplate single column demension index error #226
        /// </summary>
        [Fact]
        public async Task Issue226()
        {
            var path = PathHelper.GetTempPath();
            var templatePath = PathHelper.GetSamplePath("xlsx/TestIssue226.xlsx");
            await MiniExcel.SaveAsByTemplateAsync(path, templatePath, new { employees = new[] { new { name = "123" }, new { name = "123" } } });
            Assert.Equal("A1:A3", Helpers.GetFirstSheetDimensionRefValue(path));
        }

        /// <summary>
        /// ASP.NET Webform gridview datasource can't use miniexcel queryasdatatable · Issue #223]
        /// (https://github.com/shps951023/MiniExcel/issues/223)
        /// </summary>
        [Fact]
        public async Task Issue223()
        {
            var value = new List<Dictionary<string, object>>()
            {
                new Dictionary<string, object>(){{"A",null},{"B",null}},
                new Dictionary<string, object>(){{"A",123},{"B",new DateTime(2021,1,1)}},
                new Dictionary<string, object>(){{"A",Guid.NewGuid()},{"B","HelloWorld"}},
            };
            var path = PathHelper.GetTempPath();
            MiniExcel.SaveAs(path, value);

            var dt = await MiniExcel.QueryAsDataTableAsync(path);
            var columns = dt.Columns;
            Assert.Equal(typeof(object), columns[0].DataType);
            Assert.Equal(typeof(object), columns[1].DataType);

            Assert.Equal((double)123, dt.Rows[1]["A"]);
            Assert.Equal("HelloWorld", dt.Rows[2]["B"]);
        }

        /// <summary>
        /// [Custom yyyy-MM-dd format not convert datetime · Issue #222]
        /// (https://github.com/shps951023/MiniExcel/issues/222)
        /// </summary>
        [Fact]
        public async Task Issue222()
        {
            var path = PathHelper.GetSamplePath("xlsx/TestIssue222.xlsx");
            var q = await MiniExcel.QueryAsync(path);
            var rows = q.ToList();
            Assert.Equal(typeof(DateTime), rows[1].A.GetType());
            Assert.Equal(new DateTime(2021, 4, 29), rows[1].A);
        }

        /// <summary>
        /// QueryAsync Support StartCell #147
        /// https://github.com/shps951023/MiniExcel/issues/147
        /// </summary>
        [Fact]
        public async Task Issue147()
        {
            {
                var path = PathHelper.GetSamplePath("xlsx/TestIssue147.xlsx");
                var q = await MiniExcel.QueryAsync(path, useHeaderRow: false, startCell: "C3", sheetName: "Sheet1");
                var rows = q.ToList();
                Assert.Equal(new[] { "C", "D", "E" }, (rows[0] as IDictionary<string, object>).Keys);
                Assert.Equal(new[] { "Column1", "Column2", "Column3" }, new[] { rows[0].C as string, rows[0].D as string, rows[0].E as string });
                Assert.Equal(new[] { "C4", "D4", "E4" }, new[] { rows[1].C as string, rows[1].D as string, rows[1].E as string });
                Assert.Equal(new[] { "C9", "D9", "E9" }, new[] { rows[6].C as string, rows[6].D as string, rows[6].E as string });
                Assert.Equal(new[] { "C12", "D12", "E12" }, new[] { rows[9].C as string, rows[9].D as string, rows[9].E as string });
                Assert.Equal(new[] { "C13", "D13", "E13" }, new[] { rows[10].C as string, rows[10].D as string, rows[10].E as string });
                foreach (var i in new[] { 4, 5, 7, 8 })
                    Assert.Equal(new[] { default(string), default(string), default(string) }, new[] { rows[i].C as string, rows[i].D as string, rows[i].E as string });

                Assert.Equal(11, rows.Count);


                var columns = MiniExcel.GetColumns(path, startCell: "C3");
                Assert.Equal(new[] { "C", "D", "E" }, columns);
            }

            {
                var path = PathHelper.GetSamplePath("xlsx/TestIssue147.xlsx");
                var q = await MiniExcel.QueryAsync(path, useHeaderRow: true, startCell: "C3", sheetName: "Sheet1");
                var rows = q.ToList();
                Assert.Equal(new[] { "Column1", "Column2", "Column3" }, (rows[0] as IDictionary<string, object>).Keys);
                Assert.Equal(new[] { "C4", "D4", "E4" }, new[] { rows[0].Column1 as string, rows[0].Column2 as string, rows[0].Column3 as string });
                Assert.Equal(new[] { "C9", "D9", "E9" }, new[] { rows[5].Column1 as string, rows[5].Column2 as string, rows[5].Column3 as string });
                Assert.Equal(new[] { "C12", "D12", "E12" }, new[] { rows[8].Column1 as string, rows[8].Column2 as string, rows[8].Column3 as string });
                Assert.Equal(new[] { "C13", "D13", "E13" }, new[] { rows[9].Column1 as string, rows[9].Column2 as string, rows[9].Column3 as string });
                foreach (var i in new[] { 3, 4, 6, 7 })
                    Assert.Equal(new[] { default(string), default(string), default(string) }, new[] { rows[i].Column1 as string, rows[i].Column2 as string, rows[i].Column3 as string });

                Assert.Equal(10, rows.Count);


                var columns = MiniExcel.GetColumns(path, useHeaderRow: true, startCell: "C3");
                Assert.Equal(new[] { "Column1", "Column2", "Column3" }, columns);
            }
        }


        /// <summary>
        /// [Can SaveAs support iDataReader export to avoid the dataTable consuming too much memory · Issue #211 · shps951023/MiniExcel]
        /// (https://github.com/shps951023/MiniExcel/issues/211)
        /// </summary>
        [Fact]
        public async Task Issue211()
        {
            var path = PathHelper.GetTempPath();
            var tempSqlitePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.db");
            var connectionString = $"Data Source={tempSqlitePath};Version=3;";

            using (var connection = new SQLiteConnection(connectionString))
            {
                var reader = connection.ExecuteReader(@"select 1 Test1,2 Test2 union all select 3 , 4 union all select 5 ,6");

                MiniExcel.SaveAs(path, reader);

                var q = await MiniExcel.QueryAsync(path, true);
                var rows = q.ToList();
                Assert.Equal((double)1, rows[0].Test1);
                Assert.Equal((double)2, rows[0].Test2);
                Assert.Equal((double)3, rows[1].Test1);
                Assert.Equal((double)4, rows[1].Test2);
            }
        }

        /// <summary>
        /// [When reading Excel, can return IDataReader and DataTable to facilitate the import of database. Like ExcelDataReader provide reader.AsDataSet() · Issue #216 · shps951023/MiniExcel](https://github.com/shps951023/MiniExcel/issues/216)
        /// </summary>
        [Fact]
        public async Task Issue216()
        {
            var path = PathHelper.GetTempPath();
            var value = new[] { new { Test1 = "1", Test2 = 2 }, new { Test1 = "3", Test2 = 4 } };
            MiniExcel.SaveAs(path, value);

            {
                var table = await MiniExcel.QueryAsDataTableAsync(path);
                var columns = table.Columns;
                Assert.Equal("Test1", table.Columns[0].ColumnName);
                Assert.Equal("Test2", table.Columns[1].ColumnName);
                Assert.Equal("1", table.Rows[0]["Test1"]);
                Assert.Equal((double)2, table.Rows[0]["Test2"]);
                Assert.Equal("3", table.Rows[1]["Test1"]);
                Assert.Equal((double)4, table.Rows[1]["Test2"]);
            }

            {
                var dt = await MiniExcel.QueryAsDataTableAsync(path, false);
                Assert.Equal("Test1", dt.Rows[0]["A"]);
                Assert.Equal("Test2", dt.Rows[0]["B"]);
                Assert.Equal("1", dt.Rows[1]["A"]);
                Assert.Equal((double)2, dt.Rows[1]["B"]);
                Assert.Equal("3", dt.Rows[2]["A"]);
                Assert.Equal((double)4, dt.Rows[2]["B"]);
            }
        }

        /// <summary>
        /// https://gitee.com/dotnetchina/MiniExcel/issues/I3OSKV
        /// When exporting, the pure numeric string will be forcibly converted to a numeric type, resulting in the loss of the end data
        /// </summary>
        [Fact]
        public async Task IssueI3OSKV()
        {
            {
                var path = PathHelper.GetTempPath();
                var value = new[] { new { Test = "12345678901234567890" } };
                await MiniExcel.SaveAsAsync(path, value);

                var q = await MiniExcel.QueryAsync(path, true);
                var A2 = q.First().Test;
                Assert.Equal("12345678901234567890", A2);

                File.Delete(path);
            }

            {
                var path = PathHelper.GetTempPath();
                var value = new[] { new { Test = 123456.789 } };
                await MiniExcel.SaveAsAsync(path, value);

                var q = await MiniExcel.QueryAsync(path, true);
                var A2 = q.First().Test;
                Assert.Equal(123456.789, A2);

                File.Delete(path);
            }
        }


        /// <summary>
        /// [Dynamic QueryAsync can't summary numeric cell value default, need to cast · Issue #220 · shps951023/MiniExcel]
        /// (https://github.com/shps951023/MiniExcel/issues/220)
        /// </summary>
        [Fact]
        public async Task Issue220()
        {
            var path = PathHelper.GetSamplePath("xlsx/TestIssue220.xlsx");
            var rows = await MiniExcel.QueryAsync(path, useHeaderRow: true);
            var result = (from s in rows
                          group s by s.PRT_ID into g
                          select new
                          {
                              PRT_ID = g.Key,
                              Apr = g.Sum(_ => (double?)_.Apr),
                              May = g.Sum(_ => (double?)_.May),
                              Jun = g.Sum(_ => (double?)_.Jun),
                          }
            ).ToList();
            Assert.Equal(91843.25, result[0].Jun);
            Assert.Equal(50000.99, result[1].Jun);
        }

        /// <summary>
        /// Optimize stream excel type check 
        /// https://github.com/shps951023/MiniExcel/issues/215
        /// </summary>
        [Fact]
        public async Task Issue215()
        {
            using (var stream = new MemoryStream())
            {
                stream.SaveAs(new[] { new { V = "test1" }, new { V = "test2" } });
                var q = await stream.QueryAsync(true);
                var rows = q.ToList();
                Assert.Equal("test1", rows[0]["V"]);
                Assert.Equal("test2", rows[1]["V"]);
            }
        }

        /// <summary>
        /// Support Enum Mapping
        /// https://github.com/shps951023/MiniExcel/issues/89
        /// </summary>
        [Fact]
        public async Task Issue89()
        {
            //csv
            {
                var text = @"State
OnDuty
Fired
Leave";
                var stream = new MemoryStream();
                var writer = new StreamWriter(stream);
                writer.Write(text);
                writer.Flush();
                stream.Position = 0;
                var q = await MiniExcel.QueryAsync<Issue89VO>(stream, excelType: ExcelType.CSV);
                var rows = q.ToList();
                Assert.Equal(Issue89VO.WorkState.OnDuty, rows[0].State);
                Assert.Equal(Issue89VO.WorkState.Fired, rows[1].State);
                Assert.Equal(Issue89VO.WorkState.Leave, rows[2].State);

                var outputPath = PathHelper.GetTempPath("xlsx");
                MiniExcel.SaveAs(outputPath, rows);
                var q2 = await MiniExcel.QueryAsync<Issue89VO>(outputPath);
                var rows2 = q2.ToList();
                Assert.Equal(Issue89VO.WorkState.OnDuty, rows2[0].State);
                Assert.Equal(Issue89VO.WorkState.Fired, rows2[1].State);
                Assert.Equal(Issue89VO.WorkState.Leave, rows2[2].State);
            }

            //xlsx
            {
                var path = PathHelper.GetSamplePath("xlsx/TestIssue89.xlsx");
                var q = await MiniExcel.QueryAsync<Issue89VO>(path);
                var rows = q.ToList();
                Assert.Equal(Issue89VO.WorkState.OnDuty, rows[0].State);
                Assert.Equal(Issue89VO.WorkState.Fired, rows[1].State);
                Assert.Equal(Issue89VO.WorkState.Leave, rows[2].State);

                var outputPath = PathHelper.GetTempPath();
                MiniExcel.SaveAs(outputPath, rows);
                var q1 = await  MiniExcel.QueryAsync<Issue89VO>(outputPath);
                var rows2 = q1.ToList();
                Assert.Equal(Issue89VO.WorkState.OnDuty, rows2[0].State);
                Assert.Equal(Issue89VO.WorkState.Fired, rows2[1].State);
                Assert.Equal(Issue89VO.WorkState.Leave, rows2[2].State);
            }
        }

        public class Issue89VO
        {
            public WorkState State { get; set; }

            public enum WorkState
            {
                OnDuty,
                Leave,
                Fired
            }
        }

        /// <summary>
        /// DataTable recommended to use Caption for column name first, then use columname
        /// https://github.com/shps951023/MiniExcel/issues/217
        /// </summary>
        [Fact]
        public async Task Issue217()
        {
            DataTable table = new DataTable();
            table.Columns.Add("CustomerID");
            table.Columns.Add("CustomerName").Caption = "Name";
            table.Columns.Add("CreditLimit").Caption = "Limit";
            table.Rows.Add(new object[] { 1, "Jonathan", 23.44 });
            table.Rows.Add(new object[] { 2, "Bill", 56.87 });

            // openxml
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                MiniExcel.SaveAs(path, table);

                var q = await MiniExcel.QueryAsync(path);
                var rows = q.ToList();
                Assert.Equal("Name", rows[0].B);
                Assert.Equal("Limit", rows[0].C);


                File.Delete(path);
            }

            // csv
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.csv");
                await MiniExcel.SaveAsAsync(path, table);

                var q = await MiniExcel.QueryAsync(path);
                var rows = q.ToList();
                Assert.Equal("Name", rows[0].B);
                Assert.Equal("Limit", rows[0].C);


                File.Delete(path);
            }
        }

        /// <summary>
        /// MiniExcel.SaveAs(path, table,sheetName:“Name”) ，the actual sheetName is Sheet1
        /// https://github.com/shps951023/MiniExcel/issues/212
        /// </summary>
        [Fact]
        public async Task Issue212()
        {
            var sheetName = "Demo";
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            await MiniExcel.SaveAsAsync(path, new[] { new { x = 1, y = 2 } }, sheetName: sheetName);

            var actualSheetName = MiniExcel.GetSheetNames(path).ToList()[0];

            Assert.Equal(sheetName, actualSheetName);

            File.Delete(path);
        }

        /// <summary>
        /// Version <= v0.13.1 Template merge row list rendering has no merge
        /// https://github.com/shps951023/MiniExcel/issues/207
        /// </summary>
        [Fact]
        public async Task Issue207()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var tempaltePath = @"../../../../../samples/xlsx/TestIssue207_2.xlsx";

                var value = new
                {
                    project = new[] {
                        new {name = "項目1",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                        new {name = "項目2",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                        new {name = "項目3",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                        new {name = "項目4",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    }
                };

                await MiniExcel.SaveAsByTemplateAsync(path, tempaltePath, value);

                var q = await MiniExcel.QueryAsync(path);
                var rows = q.ToList();
                Assert.Equal("項目1", rows[0].A);
                Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[0].B);
                Assert.Equal("項目2", rows[2].A);
                Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[2].B);
                Assert.Equal("項目3", rows[4].A);
                Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[4].B);
                Assert.Equal("項目4", rows[6].A);
                Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[6].B);

                Assert.Equal("Test1", rows[8].A);
                Assert.Equal("Test2", rows[8].B);
                Assert.Equal("Test3", rows[8].C);

                Assert.Equal("項目1", rows[12].A);
                Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[12].B);
                Assert.Equal("項目2", rows[13].A);
                Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[13].B);
                Assert.Equal("項目3", rows[14].A);
                Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[14].B);
                Assert.Equal("項目4", rows[15].A);
                Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[15].B);

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:C16", demension);

                File.Delete(path);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var tempaltePath = @"../../../../../samples/xlsx/TestIssue207_Template_Merge_row_list_rendering_without_merge/template.xlsx";

                var value = new
                {
                    project = new[] {
                    new {name = "項目1",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目2",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目3",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                    new {name = "項目4",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
                }
                };

                await MiniExcel.SaveAsByTemplateAsync(path, tempaltePath, value);

                var q = await MiniExcel.QueryAsync(path);
                var rows = q.ToList();
                Assert.Equal("項目1", rows[0].A);
                Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[0].C);
                Assert.Equal("項目2", rows[3].A);
                Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[3].C);
                Assert.Equal("項目3", rows[6].A);
                Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[6].C);
                Assert.Equal("項目4", rows[9].A);
                Assert.Equal("[]內容1,[]內容2,[]內容3,[]內容4,[]內容5", rows[9].C);
                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:E15", demension);

                File.Delete(path);
            }
        }

        /// <summary>
        /// https://github.com/shps951023/MiniExcel/issues/87
        /// </summary>
        [Fact]
        public async Task Issue87()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            var templatePath = @"../../../../../samples/xlsx/TestTemplateCenterEmpty.xlsx";
            var value = new
            {
                Tests = Enumerable.Range(1, 5).Select((s, i) => new { test1 = i, test2 = i })
            };
            using (var stream = File.OpenRead(templatePath))
            {
                var q = await MiniExcel.QueryAsync(templatePath);
                var rows = q.ToList();
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);
            }

            File.Delete(path);
        }

        /// <summary>
        /// https://github.com/shps951023/MiniExcel/issues/206
        /// </summary>
        [Fact]
        public async Task Issue206()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateBasicIEmumerableFill.xlsx";

                var dt = new DataTable();
                {
                    dt.Columns.Add("name");
                    dt.Columns.Add("department");
                }
                var value = new Dictionary<string, object>()
                {
                    ["employees"] = dt
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:B2", demension);

                File.Delete(path);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateBasicIEmumerableFill.xlsx";

                var dt = new DataTable();
                {
                    dt.Columns.Add("name");
                    dt.Columns.Add("department");
                    dt.Rows.Add("Jack", "HR");
                }
                var value = new Dictionary<string, object>()
                {
                    ["employees"] = dt
                };
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:B2", demension);

                File.Delete(path);
            }
        }


        /// <summary>
        /// https://github.com/shps951023/MiniExcel/issues/193
        /// </summary>
        [Fact]
        public async Task Issue193()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateComplexWithNamespacePrefix.xlsx";

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
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);



                foreach (var sheetName in MiniExcel.GetSheetNames(path))
                {
                    var q = await MiniExcel.QueryAsync(path, sheetName: sheetName);
                    var rows = q.ToList();
                    Assert.Equal(9, rows.Count);

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

                    var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                    Assert.Equal("A1:C9", demension);

                    //TODO:row can't contain xmlns
                    //![image](https://user-images.githubusercontent.com/12729184/114998840-ead44500-9ed3-11eb-8611-58afb98faed9.png)

                }

                File.Delete(path);
            }


            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                var templatePath = @"../../../../../samples/xlsx/TestTemplateComplex.xlsx";

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
                await MiniExcel.SaveAsByTemplateAsync(path, templatePath, value);

                var q = await MiniExcel.QueryAsync(path);
                var rows = q.ToList();
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

                var demension = Helpers.GetFirstSheetDimensionRefValue(path);
                Assert.Equal("A1:C9", demension);

                File.Delete(path);
            }

        }

        [Fact]
        public async Task Issue142()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                MiniExcel.SaveAs(path, new Issue142VO[] { new Issue142VO { MyProperty1 = "MyProperty1", MyProperty2 = "MyProperty2", MyProperty3 = "MyProperty3", MyProperty4 = "MyProperty4", MyProperty5 = "MyProperty5", MyProperty6 = "MyProperty6", MyProperty7 = "MyProperty7" } });

                {
                    var q = await MiniExcel.QueryAsync(path);
                    var rows = q.ToList();
                    Assert.Equal("MyProperty4", rows[0].A);
                    Assert.Equal("CustomColumnName", rows[0].B); //note
                    Assert.Equal("MyProperty5", rows[0].C);
                    Assert.Equal("MyProperty2", rows[0].D);
                    Assert.Equal("MyProperty6", rows[0].E);
                    Assert.Equal(null, rows[0].F);
                    Assert.Equal("MyProperty3", rows[0].G);

                    Assert.Equal("MyProperty4", rows[0].A);
                    Assert.Equal("CustomColumnName", rows[0].B); //note
                    Assert.Equal("MyProperty5", rows[0].C);
                    Assert.Equal("MyProperty2", rows[0].D);
                    Assert.Equal("MyProperty6", rows[0].E);
                    Assert.Equal(null, rows[0].F);
                    Assert.Equal("MyProperty3", rows[0].G);
                }

                {
                    var q = await MiniExcel.QueryAsync<Issue142VO>(path);
                    var rows = q.ToList();

                    Assert.Equal("MyProperty4", rows[0].MyProperty4);
                    Assert.Equal("MyProperty1", rows[0].MyProperty1); //note
                    Assert.Equal("MyProperty5", rows[0].MyProperty5);
                    Assert.Equal("MyProperty2", rows[0].MyProperty2);
                    Assert.Equal("MyProperty6", rows[0].MyProperty6);
                    Assert.Null(rows[0].MyProperty7);
                    Assert.Equal("MyProperty3", rows[0].MyProperty3);
                }

                File.Delete(path);

            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.csv");
                MiniExcel.SaveAs(path, new Issue142VO[] { new Issue142VO { MyProperty1 = "MyProperty1", MyProperty2 = "MyProperty2", MyProperty3 = "MyProperty3", MyProperty4 = "MyProperty4", MyProperty5 = "MyProperty5", MyProperty6 = "MyProperty6", MyProperty7 = "MyProperty7" } });
                var expected = @"MyProperty4,CustomColumnName,MyProperty5,MyProperty2,MyProperty6,,MyProperty3
MyProperty4,MyProperty1,MyProperty5,MyProperty2,MyProperty6,,MyProperty3
";
                Assert.Equal(expected, File.ReadAllText(path));

                {
                    var q = await MiniExcel.QueryAsync<Issue142VO>(path);
                    var rows = q.ToList();

                    Assert.Equal("MyProperty4", rows[0].MyProperty4);
                    Assert.Equal("MyProperty1", rows[0].MyProperty1); //note
                    Assert.Equal("MyProperty5", rows[0].MyProperty5);
                    Assert.Equal("MyProperty2", rows[0].MyProperty2);
                    Assert.Equal("MyProperty6", rows[0].MyProperty6);
                    Assert.Null(rows[0].MyProperty7);
                    Assert.Equal("MyProperty3", rows[0].MyProperty3);
                }

                File.Delete(path);
            }

            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.csv");
                var input = new Issue142VoDuplicateColumnName[] { new Issue142VoDuplicateColumnName { } };
                Assert.Throws<InvalidOperationException>(() => MiniExcel.SaveAs(path, input));
            }
        }

        [Fact]
        public async Task Issue142_Query()
        {
            {
                var path = @"../../../../../samples/xlsx/TestIssue142.xlsx";
                await Assert.ThrowsAsync<InvalidOperationException>(async () =>
                {
                    var q = await MiniExcel.QueryAsync<Issue142VoExcelColumnNameNotFound>(path);
                    q.ToList();
                });
            }

            {
                var path = @"../../../../../samples/xlsx/TestIssue142.xlsx";
                 await Assert.ThrowsAsync<ArgumentException>(async () =>
                 {
                     var q = await MiniExcel.QueryAsync<Issue142VoOverIndex>(path);
                     q.ToList();
                 });
            }

            {
                var path = @"../../../../../samples/xlsx/TestIssue142.xlsx";
                var q = await MiniExcel.QueryAsync<Issue142VO>(path);
                var rows = q.ToList();
                Assert.Equal("CustomColumnName", rows[0].MyProperty1);
                Assert.Null(rows[0].MyProperty7);
                Assert.Equal("MyProperty2", rows[0].MyProperty2);
                Assert.Equal("MyProperty103", rows[0].MyProperty3);
                Assert.Equal("MyProperty100", rows[0].MyProperty4);
                Assert.Equal("MyProperty102", rows[0].MyProperty5);
                Assert.Equal("MyProperty6", rows[0].MyProperty6);
            }

            {
                var path = @"../../../../../samples/csv/TestIssue142.csv";
                var q = await MiniExcel.QueryAsync<Issue142VO>(path);
                var rows = q.ToList();
                Assert.Equal("CustomColumnName", rows[0].MyProperty1);
                Assert.Null(rows[0].MyProperty7);
                Assert.Equal("MyProperty2", rows[0].MyProperty2);
                Assert.Equal("MyProperty103", rows[0].MyProperty3);
                Assert.Equal("MyProperty100", rows[0].MyProperty4);
                Assert.Equal("MyProperty102", rows[0].MyProperty5);
                Assert.Equal("MyProperty6", rows[0].MyProperty6);
            }
        }

        public class Issue142VO
        {
            [ExcelColumnName("CustomColumnName")]
            public string MyProperty1 { get; set; }  //index = 1
            [ExcelIgnore]
            public string MyProperty7 { get; set; } //index = null
            public string MyProperty2 { get; set; } //index = 3
            [ExcelColumnIndex(6)]
            public string MyProperty3 { get; set; } //index = 6
            [ExcelColumnIndex("A")] // equal column index 0
            public string MyProperty4 { get; set; }
            [ExcelColumnIndex(2)]
            public string MyProperty5 { get; set; } //index = 2
            public string MyProperty6 { get; set; } //index = 4
        }

        public class Issue142VoDuplicateColumnName
        {
            [ExcelColumnIndex("A")]
            public int MyProperty1 { get; set; }
            [ExcelColumnIndex("A")]
            public int MyProperty2 { get; set; }

            public int MyProperty3 { get; set; }
            [ExcelColumnIndex("B")]
            public int MyProperty4 { get; set; }
        }

        public class Issue142VoOverIndex
        {
            [ExcelColumnIndex("Z")]
            public int MyProperty1 { get; set; }
        }

        public class Issue142VoExcelColumnNameNotFound
        {
            [ExcelColumnIndex("B")]
            public int MyProperty1 { get; set; }
        }

        /// <summary>
        /// https://github.com/shps951023/MiniExcel/issues/150
        /// </summary>
        [Fact]
        public async Task Issue150()
        {
            var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            //MiniExcel.SaveAs(path, new[] { "1", "2" });
            await Assert.ThrowsAnyAsync<NotImplementedException>(async() => await MiniExcel.SaveAsAsync(path, new[] { 1, 2 }));
            File.Delete(path);
            await Assert.ThrowsAnyAsync<NotImplementedException>(async () => await MiniExcel.SaveAsAsync(path, new[] { "1", "2" }));
            File.Delete(path);
            await Assert.ThrowsAnyAsync<NotImplementedException>(async () => await MiniExcel.SaveAsAsync(path, new[] { '1', '2' }));
            File.Delete(path);
            await Assert.ThrowsAnyAsync<NotImplementedException>(async () => await MiniExcel.SaveAsAsync(path, new[] { DateTime.Now }));
            File.Delete(path);
            await Assert.ThrowsAnyAsync<NotImplementedException>(async () => await MiniExcel.SaveAsAsync(path, new[] { Guid.NewGuid() }));
            File.Delete(path);
        }

        /// <summary>
        /// https://github.com/shps951023/MiniExcel/issues/157
        /// </summary>
        [Fact]
        public async Task Issue157()
        {
            {
                var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                Console.WriteLine("==== SaveAs by strongly type ====");
                var input = JsonConvert.DeserializeObject<IEnumerable<UserAccount>>("[{\"ID\":\"78de23d2-dcb6-bd3d-ec67-c112bbc322a2\",\"Name\":\"Wade\",\"BoD\":\"2020-09-27T00:00:00\",\"Age\":5019,\"VIP\":false,\"Points\":5019.12,\"IgnoredProperty\":null},{\"ID\":\"20d3bfce-27c3-ad3e-4f70-35c81c7e8e45\",\"Name\":\"Felix\",\"BoD\":\"2020-10-25T00:00:00\",\"Age\":7028,\"VIP\":true,\"Points\":7028.46,\"IgnoredProperty\":null},{\"ID\":\"52013bf0-9aeb-48e6-e5f5-e9500afb034f\",\"Name\":\"Phelan\",\"BoD\":\"2021-10-04T00:00:00\",\"Age\":3836,\"VIP\":true,\"Points\":3835.7,\"IgnoredProperty\":null},{\"ID\":\"3b97b87c-7afe-664f-1af5-6914d313ae25\",\"Name\":\"Samuel\",\"BoD\":\"2020-06-21T00:00:00\",\"Age\":9352,\"VIP\":false,\"Points\":9351.71,\"IgnoredProperty\":null},{\"ID\":\"9a989c43-d55f-5306-0d2f-0fbafae135bb\",\"Name\":\"Raymond\",\"BoD\":\"2021-07-12T00:00:00\",\"Age\":8210,\"VIP\":true,\"Points\":8209.76,\"IgnoredProperty\":null}]");
                await MiniExcel.SaveAsAsync(path, input);

                var q = await MiniExcel.QueryAsync(path, sheetName: "Sheet1");
                var rows = q.ToList();
                Assert.Equal(6, rows.Count());
                Assert.Equal("Sheet1", MiniExcel.GetSheetNames(path).First());

                using (var p = new ExcelPackage(new FileInfo(path)))
                {
                    var ws = p.Workbook.Worksheets.First();
                    Assert.Equal("Sheet1", ws.Name);
                    Assert.Equal("Sheet1", p.Workbook.Worksheets["Sheet1"].Name);
                }
            }
            {
                var path = @"../../../../../samples/xlsx/TestIssue157.xlsx";

                {
                    var q = await MiniExcel.QueryAsync(path, sheetName: "Sheet1");
                    var rows = q.ToList();
                    Assert.Equal(6, rows.Count());
                    Assert.Equal("Sheet1", MiniExcel.GetSheetNames(path).First());
                }
                using (var p = new ExcelPackage(new FileInfo(path)))
                {
                    var ws = p.Workbook.Worksheets.First();
                    Assert.Equal("Sheet1", ws.Name);
                    Assert.Equal("Sheet1", p.Workbook.Worksheets["Sheet1"].Name);
                }

                using (var stream = File.OpenRead(path))
                {
                    var q = await MiniExcel.QueryAsync<UserAccount>(path, sheetName: "Sheet1");
                    var rows = q.ToList();
                    Assert.Equal(5, rows.Count());

                    Assert.Equal(Guid.Parse("78DE23D2-DCB6-BD3D-EC67-C112BBC322A2"), rows[0].ID);
                    Assert.Equal("Wade", rows[0].Name);
                    Assert.Equal(DateTime.ParseExact("27/09/2020", "dd/MM/yyyy", CultureInfo.InvariantCulture), rows[0].BoD);
                    Assert.False(rows[0].VIP);
                    Assert.Equal(decimal.Parse("5019.12"), rows[0].Points);
                    Assert.Equal(1, rows[0].IgnoredProperty);
                }
            }

        }

        /// <summary>
        /// https://github.com/shps951023/MiniExcel/issues/149
        /// </summary>
        [Fact]
        public async Task Issue149()
        {
            var chars = new char[] {'\u0000','\u0001','\u0002','\u0003','\u0004','\u0005','\u0006','\u0007','\u0008',
                '\u0009', //<HT>
	           '\u000A', //<LF>
	           '\u000B','\u000C',
                '\u000D', //<CR>
	           '\u000E','\u000F','\u0010','\u0011','\u0012','\u0013','\u0014','\u0015','\u0016',
                '\u0017','\u0018','\u0019','\u001A','\u001B','\u001C','\u001D','\u001E','\u001F','\u007F'
            }.Select(s => s.ToString()).ToArray();

            {
                var path = @"../../../../../samples/xlsx/TestIssue149.xlsx";
                var q = await MiniExcel.QueryAsync(path);
                var rows = q.Select(s => (string)s.A).ToList();
                for (int i = 0; i < chars.Length; i++)
                {
                    //output.WriteLine($"{i} , {chars[i]} , {rows[i]}");
                    if (i == 13)
                        continue;
                    Assert.Equal(chars[i], rows[i]);
                }
            }

            {
                string path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var input = chars.Select(s => new { Test = s.ToString() });
                await MiniExcel.SaveAsAsync(path, input);

                var q = await MiniExcel.QueryAsync(path, true);

                var rows = q.Select(s => (string)s.Test).ToList();
                for (int i = 0; i < chars.Length; i++)
                {
                    output.WriteLine($"{i} , {chars[i]} , {rows[i]}");
                    if (i == 13 || i == 9 || i == 10)
                        continue;
                    Assert.Equal(chars[i], rows[i]);
                }
            }

            {
                string path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
                var input = chars.Select(s => new { Test = s.ToString() });
                await MiniExcel.SaveAsAsync(path, input);

                var q = await MiniExcel.QueryAsync<Issue149VO>(path);
                var rows = q.Select(s => (string)s.Test).ToList();
                for (int i = 0; i < chars.Length; i++)
                {
                    output.WriteLine($"{i} , {chars[i]} , {rows[i]}");
                    if (i == 13 || i == 9 || i == 10)
                        continue;
                    Assert.Equal(chars[i], rows[i]);
                }
            }
        }

        public class Issue149VO
        {
            public string Test { get; set; }
        }

        /// <summary>
        /// https://github.com/shps951023/MiniExcel/issues/153
        /// </summary>
        [Fact]
        public async Task Issue153()
        {
            var path = @"../../../../../samples/xlsx/TestIssue153.xlsx";
            var q = await MiniExcel.QueryAsync(path, true);
            var rows = q.First() as IDictionary<string, object>;
    
            Assert.Equal(new[] { "序号", "代号", "新代号", "名称", "XXX", "部门名称", "单位", "ERP工时   (小时)A", "工时(秒) A/3600", "标准人工工时(秒)", "生产标准机器工时(秒)", "财务、标准机器工时(秒)", "更新日期", "产品机种", "备注", "最近一次修改前的标准工时(秒)", "最近一次修改前的标准机时(秒)", "备注1" }
                , rows.Keys);
        }

        /// <summary>
        /// https://github.com/shps951023/MiniExcel/issues/137
        /// </summary>
        [Fact]
        public async Task Issue137()
        {
            var path = @"../../../../../samples/xlsx/TestIssue137.xlsx";

            {
                var q = await MiniExcel.QueryAsync(path);
                var rows = q.ToList();
                var first = rows[0] as IDictionary<string, object>; //![image](https://user-images.githubusercontent.com/12729184/113266322-ba06e400-9307-11eb-9521-d36abfda75cc.png)
                Assert.Equal(new[] { "A", "B", "C", "D", "E", "F", "G", "H" }, first.Keys.ToArray());
                Assert.Equal(11, rows.Count);
                {
                    var row = rows[0] as IDictionary<string, object>;
                    Assert.Equal("比例", row["A"]);
                    Assert.Equal("商品", row["B"]);
                    Assert.Equal("滿倉口數", row["C"]);
                    Assert.Equal(" ", row["D"]);
                    Assert.Equal(" ", row["E"]);
                    Assert.Equal(" ", row["F"]);
                    Assert.Equal(Double.Parse("0"), row["G"]);
                    Assert.Equal("1為港幣 0為台幣", row["H"]);
                }
                {
                    var row = rows[1] as IDictionary<string, object>;
                    Assert.Equal(double.Parse("1"), row["A"]);
                    Assert.Equal("MTX", row["B"]);
                    Assert.Equal(double.Parse("10"), row["C"]);
                    Assert.Null(row["D"]);
                    Assert.Null(row["E"]);
                    Assert.Null(row["F"]);
                    Assert.Null(row["G"]);
                    Assert.Null(row["H"]);
                }
                {
                    var row = rows[2] as IDictionary<string, object>;
                    Assert.Equal(double.Parse("0.95"), row["A"]);
                }
            }

            // dynamic query with head
            {
                var q = await MiniExcel.QueryAsync(path, true);
                var rows = q.ToList();
                var first = rows[0] as IDictionary<string, object>; //![image](https://user-images.githubusercontent.com/12729184/113266322-ba06e400-9307-11eb-9521-d36abfda75cc.png)
                Assert.Equal(new[] { "比例", "商品", "滿倉口數", "0", "1為港幣 0為台幣" }, first.Keys.ToArray());
                Assert.Equal(10, rows.Count);
                {
                    var row = rows[0] as IDictionary<string, object>;
                    Assert.Equal(double.Parse("1"), row["比例"]);
                    Assert.Equal("MTX", row["商品"]);
                    Assert.Equal(double.Parse("10"), row["滿倉口數"]);
                    Assert.Null(row["0"]);
                    Assert.Null(row["1為港幣 0為台幣"]);
                }

                {
                    var row = rows[1] as IDictionary<string, object>;
                    Assert.Equal(double.Parse("0.95"), row["比例"]);
                }
            }

            {
                var q = await MiniExcel.QueryAsync<Issue137ExcelRow>(path);
                var rows = q.ToList();
                Assert.Equal(10, rows.Count);
                {
                    var row = rows[0];
                    Assert.Equal(1, row.比例);
                    Assert.Equal("MTX", row.商品);
                    Assert.Equal(10, row.滿倉口數);
                }

                {
                    var row = rows[1];
                    Assert.Equal(0.95, row.比例);
                }
            }
        }

        public class Issue137ExcelRow
        {
            public double? 比例 { get; set; }
            public string 商品 { get; set; }
            public int? 滿倉口數 { get; set; }
        }


        /// <summary>
        /// https://github.com/shps951023/MiniExcel/issues/138
        /// </summary>
        [Fact]
        public async Task Issue138()
        {
            var path = @"../../../../../samples/xlsx/TestIssue138.xlsx";
            {
                var q = await MiniExcel.QueryAsync(path, true);
                var rows = q.ToList();
                Assert.Equal(6, rows.Count);

                foreach (var index in new[] { 0, 2, 5 })
                {
                    Assert.Equal(1, rows[index].實單每日損益);
                    Assert.Equal(2, rows[index].程式每日損益);
                    Assert.Equal("測試商品1", rows[index].商品);
                    Assert.Equal(111.11, rows[index].滿倉口數);
                    Assert.Equal(111.11, rows[index].波段);
                    Assert.Equal(111.11, rows[index].當沖);
                }

                foreach (var index in new[] { 1, 3, 4 })
                {
                    Assert.Null(rows[index].實單每日損益);
                    Assert.Null(rows[index].程式每日損益);
                    Assert.Null(rows[index].商品);
                    Assert.Null(rows[index].滿倉口數);
                    Assert.Null(rows[index].波段);
                    Assert.Null(rows[index].當沖);
                }
            }
            {

                var q = await MiniExcel.QueryAsync<Issue138ExcelRow>(path);
                var rows = q.ToList();
                Assert.Equal(6, rows.Count);
                Assert.Equal(new DateTime(2021, 3, 1), rows[0].date);

                foreach (var index in new[] { 0, 2, 5 })
                {
                    Assert.Equal(1, rows[index].實單每日損益);
                    Assert.Equal(2, rows[index].程式每日損益);
                    Assert.Equal("測試商品1", rows[index].商品);
                    Assert.Equal(111.11, rows[index].滿倉口數);
                    Assert.Equal(111.11, rows[index].波段);
                    Assert.Equal(111.11, rows[index].當沖);
                }

                foreach (var index in new[] { 1, 3, 4 })
                {
                    Assert.Null(rows[index].實單每日損益);
                    Assert.Null(rows[index].程式每日損益);
                    Assert.Null(rows[index].商品);
                    Assert.Null(rows[index].滿倉口數);
                    Assert.Null(rows[index].波段);
                    Assert.Null(rows[index].當沖);
                }
            }
        }

        public class Issue138ExcelRow
        {
            public DateTime? date { get; set; }
            public int? 實單每日損益 { get; set; }
            public int? 程式每日損益 { get; set; }
            public string 商品 { get; set; }
            public double? 滿倉口數 { get; set; }
            public double? 波段 { get; set; }
            public double? 當沖 { get; set; }
        }
    }
}