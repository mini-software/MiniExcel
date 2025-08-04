namespace MiniExcelLib.Csv.Tests;

public class IssueTests
{
    private readonly CsvExporter _csvExporter = MiniExcel.Exporters.GetCsvExporter();
    private readonly CsvImporter _csvImporter = MiniExcel.Importers.GetCsvImporter();

    private readonly OpenXmlExporter _openXmlExporter = MiniExcel.Exporters.GetOpenXmlExporter();
    private readonly OpenXmlImporter _openXmlImporter = MiniExcel.Importers.GetOpenXmlImporter();

    [Fact]
    public void TestPR10()
    {
        var path = PathHelper.GetFile("csv/TestIssue142.csv");
        var config = new CsvConfiguration
        {
            SplitFn = row => Regex.Split(row, "[\t,](?=(?:[^\"]|\"[^\"]*\")*$)")
                .Select(s => Regex.Replace(s.Replace("\"\"", "\""), "^\"|\"$", ""))
                .ToArray()
        };
        var rows = _csvImporter.Query(path, configuration: config).ToList();
    }

    /// <summary>
    /// https://gitee.com/dotnetchina/MiniExcel/issues/I49RZH
    /// https://github.com/mini-software/MiniExcel/issues/305
    /// </summary>
    [Fact]
    public void TestIssueI49RZH()
    {
        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        var value = new[]
        {
            new TestIssueI49RZHDto{ dd = DateTimeOffset.Parse("2022-01-22")},
            new TestIssueI49RZHDto{ dd = null}
        };
        _csvExporter.Export(path.ToString(), value);

        var rows =  _csvImporter.Query(path.ToString()).ToList();
        Assert.Equal("2022-01-22", rows[1].A);
    }

    private class TestIssueI49RZHDto
    {
        [MiniExcelFormat("yyyy-MM-dd")]
        public DateTimeOffset? dd { get; set; }
    }
    
    /// <summary>
    /// https://gitee.com/dotnetchina/MiniExcel/issues/I4X92G
    /// </summary>
    [Fact]
    public void TestIssueI4X92G()
    {
        using var file = AutoDeletingPath.Create(ExcelType.Csv);
        var path = file.ToString();

        {
            var value = new[]
            {
                new { ID = 1, Name = "Jack", InDate = new DateTime(2021,01,03)},
                new { ID = 2, Name = "Henry", InDate = new DateTime(2020,05,03)}
            };
            _csvExporter.Export(path, value);
            var content = File.ReadAllText(path);
            Assert.Equal(
                """
                 ID,Name,InDate
                 1,Jack,"2021-01-03 00:00:00"
                 2,Henry,"2020-05-03 00:00:00"

                 """,
                content);
        }
        {
            var value = new { ID = 3, Name = "Mike", InDate = new DateTime(2021, 04, 23) };
            var rowsWritten = _csvExporter.Append(path, value);
            Assert.Equal(1, rowsWritten);

            var content = File.ReadAllText(path);
            Assert.Equal(
                """
                 ID,Name,InDate
                 1,Jack,"2021-01-03 00:00:00"
                 2,Henry,"2020-05-03 00:00:00"
                 3,Mike,"2021-04-23 00:00:00"

                 """,
                content);
        }
        {
            var value = new[]
            {
                new { ID=4,Name ="Frank",InDate=new DateTime(2021,06,07)},
                new { ID=5,Name ="Gloria",InDate=new DateTime(2022,05,03)},
            };
            var rowsWritten = _csvExporter.Append(path, value);
            Assert.Equal(2, rowsWritten);

            var content = File.ReadAllText(path);
            Assert.Equal(
                """
                 ID,Name,InDate
                 1,Jack,"2021-01-03 00:00:00"
                 2,Henry,"2020-05-03 00:00:00"
                 3,Mike,"2021-04-23 00:00:00"
                 4,Frank,"2021-06-07 00:00:00"
                 5,Gloria,"2022-05-03 00:00:00"

                 """,
                content);
        }
    }

    private class TestIssue316Dto
    {
        public decimal Amount { get; set; }
        public DateTime CreateTime { get; set; }
    }

    /// <summary>
    /// Using stream.SaveAs will close the Stream automatically when Specifying excelType
    /// https://gitee.com/dotnetchina/MiniExcel/issues/I57WMM
    /// </summary>
    [Fact]
    public void TestIssueI57WMM()
    {
        Dictionary<string, object>[] sheets = [new() { ["ID"] = "0001", ["Name"] = "Jack" }];
        using var stream = new MemoryStream();

        var config = new CsvConfiguration { StreamWriterFunc = x => new StreamWriter(x, Encoding.Default, leaveOpen: true) };
        _csvExporter.Export(stream, sheets, configuration: config);
        stream.Seek(0, SeekOrigin.Begin);

        // convert stream to string
        using var reader = new StreamReader(stream);
        var text = reader.ReadToEnd();

        Assert.Equal("ID,Name\r\n0001,Jack\r\n", text);
    }

    [Fact]
    public void Issue217()
    {
        using var table = new DataTable();
        table.Columns.Add("CustomerID");
        table.Columns.Add("CustomerName").Caption = "Name";
        table.Columns.Add("CreditLimit").Caption = "Limit";
        table.Rows.Add(1, "Jonathan", 23.44);
        table.Rows.Add(2, "Bill", 56.87);

        
        
        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        _csvExporter.Export(path.ToString(), table);

        var rows =  _csvImporter.Query(path.ToString()).ToList();
        Assert.Equal("Name", rows[0].B);
        Assert.Equal("Limit", rows[0].C);
    }
    
    /// <summary>
    /// Csv Query split comma not correct #237
    /// https://github.com/mini-software/MiniExcel/issues/237
    /// </summary>
    [Fact]
    public void Issue237()
    {
        var value = new[]
        {
            new{ id = "\"\"1,2,3\"\"" },
            new{ id = "1,2,3" }
        };
        
        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        _csvExporter.Export(path.ToString(), value);

        var rows =  _csvImporter.Query(path.ToString(), true).ToList();

        Assert.Equal("\"\"1,2,3\"\"", rows[0].id);
        Assert.Equal("1,2,3", rows[1].id);
    }

    /// <summary>
    /// Support Custom Datetime format #241
    /// </summary>
    [Fact]
    public void Issue241()
    {
        Issue241Dto[] value =
        [
            new() { Name = "Jack", InDate = new DateTime(2021, 01, 04) },
            new() { Name = "Henry", InDate = new DateTime(2020, 04, 05) }
        ];

        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        _csvExporter.Export(path.ToString(), value);

        var rows1 = _csvImporter.Query(path.ToString(), true).ToList();
        Assert.Equal(rows1[0].InDate, "01 04, 2021");
        Assert.Equal(rows1[1].InDate, "04 05, 2020");

        var rows2 = _csvImporter.Query<Issue241Dto>(path.ToString()).ToList();
        Assert.Equal(rows2[0].InDate, new DateTime(2021, 01, 04));
        Assert.Equal(rows2[1].InDate, new DateTime(2020, 04, 05));
    }
    
    private class Issue241Dto
    {
        public string Name { get; set; }

        [MiniExcelFormat("MM dd, yyyy")]
        public DateTime InDate { get; set; }
    }
    
    
    /// <summary>
    /// Csv type mapping Query error "cannot be converted to xxx type" #243
    /// </summary>
    [Fact]
    public void Issue243()
    {
        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        var value = new[]
        {
            new { Name = "Jack", Age = 25, InDate = new DateTime(2021,01,03) },
            new { Name = "Henry",  Age = 36, InDate = new DateTime(2020,05,03) },
        };
        _csvExporter.Export(path.ToString(), value);

        var rows =  _csvImporter.Query<Issue243Dto>(path.ToString()).ToList();
        Assert.Equal("Jack", rows[0].Name);
        Assert.Equal(25, rows[0].Age);
        Assert.Equal(new DateTime(2021, 01, 03), rows[0].InDate);

        Assert.Equal("Henry", rows[1].Name);
        Assert.Equal(36, rows[1].Age);
        Assert.Equal(new DateTime(2020, 05, 03), rows[1].InDate);
    }

    private class Issue243Dto
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public DateTime InDate { get; set; }
    }
    
    
    /// <summary>
    /// https://github.com/mini-software/MiniExcel/issues/312
    /// </summary>
    [Fact]
    public void TestIssue312()
    {
        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        TestIssue312Dto[] value =
        [
            new() { Value = 12345.6789},
            new() { Value = null}
        ];
        _csvExporter.Export(path.ToString(), value);

        var rows =  _csvImporter.Query(path.ToString()).ToList();
        Assert.Equal("12,345.68", rows[1].A);
    }

    private class TestIssue312Dto
    {
        [MiniExcelFormat("0,0.00")]
        public double? Value { get; set; }
    }
    
    [Fact]
    public async Task TestIssue338()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        {
            var path = PathHelper.GetFile("csv/TestIssue338.csv");
            var row = _csvImporter.QueryAsync(path).ToBlockingEnumerable().FirstOrDefault();
            Assert.Equal("���Ĳ�������", row!.A);
        }
        {
            var path = PathHelper.GetFile("csv/TestIssue338.csv");
            var config = new CsvConfiguration
            {
                StreamReaderFunc = stream => new StreamReader(stream, Encoding.GetEncoding("gb2312"))
            };
            var row = _csvImporter.QueryAsync(path, configuration: config).ToBlockingEnumerable().FirstOrDefault();
            Assert.Equal("中文测试内容", row!.A);
        }
        {
            var path = PathHelper.GetFile("csv/TestIssue338.csv");
            var config = new CsvConfiguration
            {
                StreamReaderFunc = stream => new StreamReader(stream, Encoding.GetEncoding("gb2312"))
            };
            await using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                var row = _csvImporter.QueryAsync(stream, configuration: config).ToBlockingEnumerable().FirstOrDefault();
                Assert.Equal("中文测试内容", row!.A);
            }
        }
    }

    [Fact]
    public void TestIssueI4WDA9()
    {
        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        var value = new DataTable();
        {
            value.Columns.Add("\"name\"");
            value.Rows.Add("\"Jack\"");
        }

        _csvExporter.Export(path.ToString(), value);
        Assert.Equal("\"\"\"name\"\"\"\r\n\"\"\"Jack\"\"\"\r\n", File.ReadAllText(path.ToString()));
    }

    [Fact]
    public void TestIssue316()
    {
        // XLSX
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            var value = new[]
            {
                new{ Amount=123_456.789M, CreateTime=DateTime.Parse("2018-01-31",CultureInfo.InvariantCulture)}
            };
            var config = new OpenXmlConfiguration
            {
                Culture = new CultureInfo("fr-FR"),
            };
            _openXmlExporter.Export(path, value, configuration: config);

            //Datetime error
            Assert.Throws<MiniExcelInvalidCastException>(() =>
            {
                var conf = new OpenXmlConfiguration
                {
                    Culture = new CultureInfo("en-US"),
                };
                _ = _openXmlImporter.Query<TestIssue316Dto>(path, configuration: conf).ToList();
            });

            // dynamic
            var rows = _openXmlImporter.Query(path, true).ToList();
            Assert.Equal("123456,789", rows[0].Amount);
            Assert.Equal("31/01/2018 00:00:00", rows[0].CreateTime);
        }

        // type
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            var value = new[]
            {
                new { Amount = 123_456.789M, CreateTime = new DateTime(2018, 5, 12) }
            };
            {
                var config = new OpenXmlConfiguration
                {
                    Culture = new CultureInfo("fr-FR"),
                };
                _openXmlExporter.Export(path, value, configuration: config);
            }

            {
                var rows = _openXmlImporter.Query(path, true).ToList();
                Assert.Equal("123456,789", rows[0].Amount);
                Assert.Equal("12/05/2018 00:00:00", rows[0].CreateTime);
            }

            {
                var config = new OpenXmlConfiguration
                {
                    Culture = new CultureInfo("en-US"),
                };
                var rows = _openXmlImporter.Query<TestIssue316Dto>(path, configuration: config).ToList();

                Assert.Equal("2018-12-05 00:00:00", rows[0].CreateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                Assert.Equal(123456789m, rows[0].Amount);
            }

            {
                var config = new OpenXmlConfiguration
                {
                    Culture = new CultureInfo("fr-FR"),
                };
                var rows = _openXmlImporter.Query<TestIssue316Dto>(path, configuration: config).ToList();

                Assert.Equal("2018-05-12 00:00:00", rows[0].CreateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                Assert.Equal(123456.789m, rows[0].Amount);
            }
        }

        // CSV
        {
            using var file = AutoDeletingPath.Create(ExcelType.Csv);
            var path = file.ToString();
            var value = new[]
            {
                new { Amount = 123_456.789M, CreateTime = new DateTime(2018, 1, 31) }
            };

            var config = new CsvConfiguration
            {
                Culture = new CultureInfo("fr-FR"),
            };
            _csvExporter.Export(path, value, configuration: config);

            //Datetime error
            Assert.Throws<MiniExcelInvalidCastException>(() =>
            {
                var conf = new CsvConfiguration
                {
                    Culture = new CultureInfo("en-US")
                };
                _ = _csvImporter.Query<TestIssue316Dto>(path, configuration: conf).ToList();
            });

            // dynamic
            var rows = _csvImporter.Query(path, true).ToList();
            Assert.Equal("123456,789", rows[0].Amount);
            Assert.Equal("31/01/2018 00:00:00", rows[0].CreateTime);
        }

        // type
        {
            using var file = AutoDeletingPath.Create(ExcelType.Csv);
            var path = file.ToString();

            var value = new[]
            {
                new{ Amount=123_456.789M, CreateTime=DateTime.Parse("2018-05-12", CultureInfo.InvariantCulture)}
            };
            {
                var config = new CsvConfiguration
                {
                    Culture = new CultureInfo("fr-FR"),
                };
                _csvExporter.Export(path, value, configuration: config);
            }

            {
                var rows = _csvImporter.Query(path, true).ToList();
                Assert.Equal("123456,789", rows[0].Amount);
                Assert.Equal("12/05/2018 00:00:00", rows[0].CreateTime);
            }

            {
                var config = new CsvConfiguration
                {
                    Culture = new CultureInfo("en-US"),
                };
                var rows = _csvImporter.Query<TestIssue316Dto>(path, configuration: config).ToList();

                Assert.Equal("2018-12-05 00:00:00", rows[0].CreateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                Assert.Equal(123456789m, rows[0].Amount);
            }

            {
                var config = new CsvConfiguration
                {
                    Culture = new CultureInfo("fr-FR"),
                };
                var rows = _csvImporter.Query<TestIssue316Dto>(path, configuration: config).ToList();

                Assert.Equal("2018-05-12 00:00:00", rows[0].CreateTime.ToString("yyyy-MM-dd HH:mm:ss"));
                Assert.Equal(123456.789m, rows[0].Amount);
            }
        }
    }

    /// <summary>
    /// Column '' does not belong to table when csv convert to datatable #298
    /// https://github.com/mini-software/MiniExcel/issues/298
    /// </summary>
    [Fact]
    public void TestIssue298()
    {
        var path = PathHelper.GetFile("/csv/TestIssue298.csv");
#pragma warning disable CS0618 // Type or member is obsolete
        var dt = _csvImporter.QueryAsDataTable(path);
#pragma warning restore CS0618
        Assert.Equal(["ID", "Name", "Age"], dt.Columns.Cast<DataColumn>().Select(x => x.ColumnName));
    }
    /// <summary>
    /// [According to the XLSX to CSV example, there will be data loss if there is no header. · Issue #292 · mini-software/MiniExcel](https://github.com/mini-software/MiniExcel/issues/292)
    /// </summary>
    [Fact]
    public void TestIssue292()
    {
        {
            var xlsxPath = PathHelper.GetFile("/xlsx/TestIssue292.xlsx");
            using var csvPath = AutoDeletingPath.Create(ExcelType.Csv);
            _csvExporter.ConvertXlsxToCsv(xlsxPath, csvPath.ToString(), false);

            var actualCotent = File.ReadAllText(csvPath.ToString());
            Assert.Equal(
                """
                Name,Age,Name,Age
                Jack,22,Mike,25
                Henry,44,Jerry,44

                """,
                actualCotent);
        }

        {
            var csvPath = PathHelper.GetFile("/csv/TestIssue292.csv");
            using var path = AutoDeletingPath.Create();
            _csvExporter.ConvertCsvToXlsx(csvPath, path.ToString());

            var rows = _openXmlImporter.Query(path.ToString()).ToList();
            Assert.Equal(3, rows.Count);
            Assert.Equal("Name", rows[0].A);
            Assert.Equal("Age", rows[0].B);
            Assert.Equal("Name", rows[0].C);
            Assert.Equal("Age", rows[0].D);
            Assert.Equal("Jack", rows[1].A);
            Assert.Equal("22", rows[1].B);
            Assert.Equal("Mike", rows[1].C);
            Assert.Equal("25", rows[1].D);
        }
    }

    /// <summary>
    /// [Csv Query then SaveAs will throw "Stream was not readable." exception · Issue #293 · mini-software/MiniExcel](https://github.com/mini-software/MiniExcel/issues/293)
    /// </summary>
    [Fact]
    public void TestIssue293()
    {
        var path = PathHelper.GetFile("/csv/Test5x2.csv");
        using var tempPath = AutoDeletingPath.Create();
        using var csv = File.OpenRead(path);
        var value = _csvImporter.Query(csv, useHeaderRow: false);
        _openXmlExporter.Export(tempPath.ToString(), value, printHeader: false);
    }

    /// <summary>
    /// Csv not support QueryAsDataTable #279 https://github.com/mini-software/MiniExcel/issues/279
    /// </summary>
    [Fact]
    public void TestIssue279()
    {
        var path = PathHelper.GetFile("/csv/TestHeader.csv");
#pragma warning disable CS0618 // Type or member is obsolete
        using var dt = _csvImporter.QueryAsDataTable(path);
#pragma warning restore CS0618
        Assert.Equal("A1", dt.Rows[0]["Column1"]);
        Assert.Equal("A2", dt.Rows[1]["Column1"]);
        Assert.Equal("B1", dt.Rows[0]["Column2"]);
        Assert.Equal("B2", dt.Rows[1]["Column2"]);
    }

    /// <summary>
    /// [Convert csv to xlsx · Issue #261 · mini-software/MiniExcel](https://github.com/mini-software/MiniExcel/issues/261)
    /// </summary>
    [Fact]
    public void TestIssue261()
    {
        var csvPath = PathHelper.GetFile("csv/TestCsvToXlsx.csv");
        using var path = AutoDeletingPath.Create();

        _csvExporter.ConvertCsvToXlsx(csvPath, path.FilePath);
        var rows = _openXmlImporter.Query(path.ToString()).ToList();

        Assert.Equal("Name", rows[0].A);
        Assert.Equal("Jack", rows[1].A);
        Assert.Equal("Neo", rows[2].A);
        Assert.Null(rows[3].A);
        Assert.Null(rows[4].A);
        Assert.Equal("Age", rows[0].B);
        Assert.Equal("34", rows[1].B);
        Assert.Equal("26", rows[2].B);
        Assert.Null(rows[3].B);
        Assert.Null(rows[4].B);
    }

    /// <summary>
    /// Csv SaveAs by datareader with encoding default show messy code #253
    /// </summary>
    [Fact]
    public void Issue253()
    {
        {
            var value = new[] { new { col1 = "世界你好" } };
            using var path = AutoDeletingPath.Create(ExcelType.Csv);
            _csvExporter.Export(path.ToString(), value);
            const string expected =
                """
                col1
                世界你好

                """;
            Assert.Equal(expected, File.ReadAllText(path.ToString()));
        }

        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var value = new[] { new { col1 = "世界你好" } };
            using var path = AutoDeletingPath.Create(ExcelType.Csv);
            var config = new CsvConfiguration
            {
                StreamWriterFunc = stream => new StreamWriter(stream, Encoding.GetEncoding("gb2312"))
            };
            _csvExporter.Export(path.ToString(), value, configuration: config);
            const string expected =
                """
                col1
                �������

                """;
            Assert.Equal(expected, File.ReadAllText(path.ToString()));
        }

        using var cn = Db.GetConnection();

        {
            var value = cn.ExecuteReader("select '世界你好' col1");
            using var path = AutoDeletingPath.Create(ExcelType.Csv);
            _csvExporter.Export(path.ToString(), value);
            const string expected =
                """
                col1
                世界你好

                """;
            Assert.Equal(expected, File.ReadAllText(path.ToString()));
        }
    }

    /// <summary>
    /// [CSV SaveAs support datareader · Issue #251 · mini-software/MiniExcel](https://github.com/mini-software/MiniExcel/issues/251)
    /// </summary>
    [Fact]
    public void Issue251()
    {
        using var cn = Db.GetConnection();
        using var reader = cn.ExecuteReader(@"select '""<>+-*//}{\\n' a,1234567890 b union all select '<test>Hello World</test>',-1234567890");
        using var path = AutoDeletingPath.Create(ExcelType.Csv);
        _csvExporter.Export(path.ToString(), reader);
        const string expected =
            """"
            a,b
            """<>+-*//}{\\n",1234567890
            "<test>Hello World</test>",-1234567890

            """";

        Assert.Equal(expected, File.ReadAllText(path.ToString()));
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
    /// Support Enum Mapping
    /// https://github.com/mini-software/MiniExcel/issues/89
    /// </summary>
    [Fact]
    public void Issue89()
    {
        //csv
        {
            const string text =
               """
               State
               OnDuty
               Fired
               Leave
               """;

            using var stream = new MemoryStream();
            using var writer = new StreamWriter(stream);

            writer.Write(text);
            writer.Flush();
            stream.Position = 0;
            var rows = _csvImporter.Query(stream, useHeaderRow: true).ToList();

            Assert.Equal(nameof(Issue89VO.WorkState.OnDuty), rows[0].State);
            Assert.Equal(nameof(Issue89VO.WorkState.Fired), rows[1].State);
            Assert.Equal(nameof(Issue89VO.WorkState.Leave), rows[2].State);

            using var path = AutoDeletingPath.Create(ExcelType.Csv);
            _csvExporter.Export(path.ToString(), rows);
            var rows2 = _csvImporter.Query<Issue89VO>(path.ToString()).ToList();

            Assert.Equal(Issue89VO.WorkState.OnDuty, rows2[0].State);
            Assert.Equal(Issue89VO.WorkState.Fired, rows2[1].State);
            Assert.Equal(Issue89VO.WorkState.Leave, rows2[2].State);
        }

        //xlsx
        {
            var path = PathHelper.GetFile("xlsx/TestIssue89.xlsx");
            var rows = _openXmlImporter.Query<Issue89VO>(path).ToList();

            Assert.Equal(Issue89VO.WorkState.OnDuty, rows[0].State);
            Assert.Equal(Issue89VO.WorkState.Fired, rows[1].State);
            Assert.Equal(Issue89VO.WorkState.Leave, rows[2].State);

            using var xlsxPath = AutoDeletingPath.Create();
            _openXmlExporter.Export(xlsxPath.ToString(), rows);
            var rows2 = _openXmlImporter.Query<Issue89VO>(xlsxPath.ToString()).ToList();

            Assert.Equal(Issue89VO.WorkState.OnDuty, rows2[0].State);
            Assert.Equal(Issue89VO.WorkState.Fired, rows2[1].State);
            Assert.Equal(Issue89VO.WorkState.Leave, rows2[2].State);
        }
    }

    private class Issue142VoDuplicateColumnName
    {
        [MiniExcelColumnIndex("A")]
        public int MyProperty1 { get; set; }
        [MiniExcelColumnIndex("A")]
        public int MyProperty2 { get; set; }

        public int MyProperty3 { get; set; }
        [MiniExcelColumnIndex("B")]
        public int MyProperty4 { get; set; }
    }

    private class Issue142VO
    {
        [MiniExcelColumnName("CustomColumnName")]
        public string MyProperty1 { get; set; }  //index = 1
        [MiniExcelIgnore]
        public string MyProperty7 { get; set; } //index = null
        public string MyProperty2 { get; set; } //index = 3
        [MiniExcelColumnIndex(6)]
        public string MyProperty3 { get; set; } //index = 6
        [MiniExcelColumnIndex("A")] // equal column index 0
        public string MyProperty4 { get; set; }
        [MiniExcelColumnIndex(2)]
        public string MyProperty5 { get; set; } //index = 2
        public string MyProperty6 { get; set; } //index = 4
    }

    [Fact]
    public void Issue142()
    {
        {
            using var file = AutoDeletingPath.Create();
            var path = file.ToString();
            Issue142VO[] values =
            [
                new()
                {
                    MyProperty1 = "MyProperty1", MyProperty2 = "MyProperty2", MyProperty3 = "MyProperty3",
                    MyProperty4 = "MyProperty4", MyProperty5 = "MyProperty5", MyProperty6 = "MyProperty6",
                    MyProperty7 = "MyProperty7"
                }
            ];
            var rowsWritten = _openXmlExporter.Export(path, values);
            Assert.Single(rowsWritten);
            Assert.Equal(1, rowsWritten[0]);

            {
                var rows = _openXmlImporter.Query(path).ToList();

                Assert.Equal("MyProperty4", rows[0].A);
                Assert.Equal("CustomColumnName", rows[0].B);
                Assert.Equal("MyProperty5", rows[0].C);
                Assert.Equal("MyProperty2", rows[0].D);
                Assert.Equal("MyProperty6", rows[0].E);
                Assert.Equal(null, rows[0].F);
                Assert.Equal("MyProperty3", rows[0].G);

                Assert.Equal("MyProperty4", rows[0].A);
                Assert.Equal("CustomColumnName", rows[0].B);
                Assert.Equal("MyProperty5", rows[0].C);
                Assert.Equal("MyProperty2", rows[0].D);
                Assert.Equal("MyProperty6", rows[0].E);
                Assert.Equal(null, rows[0].F);
                Assert.Equal("MyProperty3", rows[0].G);
            }

            {
                var rows = _openXmlImporter.Query<Issue142VO>(path).ToList();

                Assert.Equal("MyProperty4", rows[0].MyProperty4);
                Assert.Equal("MyProperty1", rows[0].MyProperty1);
                Assert.Equal("MyProperty5", rows[0].MyProperty5);
                Assert.Equal("MyProperty2", rows[0].MyProperty2);
                Assert.Equal("MyProperty6", rows[0].MyProperty6);
                Assert.Null(rows[0].MyProperty7);
                Assert.Equal("MyProperty3", rows[0].MyProperty3);
            }
        }

        {
            using var file = AutoDeletingPath.Create(ExcelType.Csv);
            var path = file.ToString();
            Issue142VO[] values =
            [
                new()
                {
                    MyProperty1 = "MyProperty1", MyProperty2 = "MyProperty2", MyProperty3 = "MyProperty3",
                    MyProperty4 = "MyProperty4", MyProperty5 = "MyProperty5", MyProperty6 = "MyProperty6",
                    MyProperty7 = "MyProperty7"
                }
            ];
            var rowsWritten = _csvExporter.Export(path, values);
            Assert.Single(rowsWritten);
            Assert.Equal(1, rowsWritten[0]);

            const string expected =
                """
                MyProperty4,CustomColumnName,MyProperty5,MyProperty2,MyProperty6,,MyProperty3
                MyProperty4,MyProperty1,MyProperty5,MyProperty2,MyProperty6,,MyProperty3

                """;

            Assert.Equal(expected, File.ReadAllText(path));

            {
                var rows = _csvImporter.Query<Issue142VO>(path).ToList();

                Assert.Equal("MyProperty4", rows[0].MyProperty4);
                Assert.Equal("MyProperty1", rows[0].MyProperty1);
                Assert.Equal("MyProperty5", rows[0].MyProperty5);
                Assert.Equal("MyProperty2", rows[0].MyProperty2);
                Assert.Equal("MyProperty6", rows[0].MyProperty6);
                Assert.Null(rows[0].MyProperty7);
                Assert.Equal("MyProperty3", rows[0].MyProperty3);
            }
        }

        {
            using var path = AutoDeletingPath.Create(ExcelType.Csv);
            Issue142VoDuplicateColumnName[] input =
            [
                new() { MyProperty1 = 0, MyProperty2 = 0, MyProperty3 = 0, MyProperty4 = 0 }
            ];
            Assert.Throws<InvalidOperationException>(() => _csvExporter.Export(path.ToString(), input));
        }
    }

    [Fact]
    public void Issue142_Query()
    {
        const string path = "../../../../../samples/xlsx/TestIssue142.xlsx";
        const string csvPath = "../../../../../samples/csv/TestIssue142.csv";
        {
            var rows = _openXmlImporter.Query<Issue142VoExcelColumnNameNotFound>(path).ToList();
            Assert.Equal(0, rows[0].MyProperty1);
        }

        Assert.Throws<ArgumentException>(() => _openXmlImporter.Query<Issue142VoOverIndex>(path).ToList());

        var rowsXlsx = _openXmlImporter.Query<Issue142VO>(path).ToList();
        Assert.Equal("CustomColumnName", rowsXlsx[0].MyProperty1);
        Assert.Null(rowsXlsx[0].MyProperty7);
        Assert.Equal("MyProperty2", rowsXlsx[0].MyProperty2);
        Assert.Equal("MyProperty103", rowsXlsx[0].MyProperty3);
        Assert.Equal("MyProperty100", rowsXlsx[0].MyProperty4);
        Assert.Equal("MyProperty102", rowsXlsx[0].MyProperty5);
        Assert.Equal("MyProperty6", rowsXlsx[0].MyProperty6);

        var rowsCsv = _csvImporter.Query<Issue142VO>(csvPath).ToList();
        Assert.Equal("CustomColumnName", rowsCsv[0].MyProperty1);
        Assert.Null(rowsCsv[0].MyProperty7);
        Assert.Equal("MyProperty2", rowsCsv[0].MyProperty2);
        Assert.Equal("MyProperty103", rowsCsv[0].MyProperty3);
        Assert.Equal("MyProperty100", rowsCsv[0].MyProperty4);
        Assert.Equal("MyProperty102", rowsCsv[0].MyProperty5);
        Assert.Equal("MyProperty6", rowsCsv[0].MyProperty6);
    }

    private class Issue142VoOverIndex
    {
        [MiniExcelColumnIndex("Z")]
        public int MyProperty1 { get; set; }
    }

    private class Issue142VoExcelColumnNameNotFound
    {
        [MiniExcelColumnIndex("B")]
        public int MyProperty1 { get; set; }
    }

    private class Issue507V01
    {
        public string A { get; set; }
        public DateTime B { get; set; }
        public string C { get; set; }
        public int D { get; set; }
    }


    [Fact]
    public void Issue507_1()
    {
        //Problem with multi-line when using Query func
        //https://github.com/mini-software/MiniExcel/issues/507

        var path = Path.Combine(Path.GetTempPath(), string.Concat(nameof(IssueTests), "_", nameof(Issue507_1), ".csv"));
        var values = new Issue507V01[]
        {
            new() { A = "Github", B = DateTime.Parse("2021-01-01"), C = "abcd", D = 123 },
            new() { A = "Microsoft \nTest 1", B = DateTime.Parse("2021-02-01"), C = "efgh", D = 123 },
            new() { A = "Microsoft \rTest 2", B = DateTime.Parse("2021-02-01"), C = "ab\nc\nd", D = 123 },
            new() { A = "Microsoft\"\" \r\nTest\n3", B = DateTime.Parse("2021-02-01"), C = "a\"\"\nb\n\nc", D = 123 },
        };

        var config = new CsvConfiguration
        {
            //AlwaysQuote = true,
            ReadLineBreaksWithinQuotes = true,
        };

        // create
        using (var stream = File.Create(path))
        {
            _csvExporter.Export(stream, values, configuration: config);
        }

        // read
        var getRowsInfo = _csvImporter.Query<Issue507V01>(path, configuration: config).ToArray();

        Assert.Equal(values.Length, getRowsInfo.Length);

        Assert.Equal("Github", getRowsInfo[0].A);
        Assert.Equal("abcd", getRowsInfo[0].C);

        Assert.Equal($"Microsoft {config.NewLine}Test 1", getRowsInfo[1].A);
        Assert.Equal("efgh", getRowsInfo[1].C);

        Assert.Equal($"Microsoft {config.NewLine}Test 2", getRowsInfo[2].A);
        Assert.Equal($"ab{config.NewLine}c{config.NewLine}d", getRowsInfo[2].C);

        Assert.Equal($"""Microsoft"" {config.NewLine}Test{config.NewLine}3""", getRowsInfo[3].A);
        Assert.Equal($"""a""{config.NewLine}b{config.NewLine}{config.NewLine}c""", getRowsInfo[3].C);

        File.Delete(path);
    }

    private class Issue507V02
    {
        public DateTime B { get; set; }
        public int D { get; set; }
    }

    [Fact]
    public void Issue507_2()
    {
        //Problem with multi-line when using Query func
        //https://github.com/mini-software/MiniExcel/issues/507

        var path = Path.Combine(Path.GetTempPath(), string.Concat(nameof(IssueTests), "_", nameof(Issue507_2), ".csv"));
        var values = new Issue507V02[]
        {
            new() { B = DateTime.Parse("2021-01-01"), D = 123 },
            new() { B = DateTime.Parse("2021-02-01"), D = 123 },
            new() { B = DateTime.Parse("2021-02-01"), D = 123 },
            new() { B = DateTime.Parse("2021-02-01"), D = 123 },
        };

        var config = new CsvConfiguration
        {
            //AlwaysQuote = true,
            ReadLineBreaksWithinQuotes = true,
        };

        // create
        using (var stream = File.Create(path))
        {
            _csvExporter.Export(stream, values, true, config);
        }

        // read
        var getRowsInfo = _csvImporter.Query<Issue507V02>(path, configuration: config).ToArray();
        Assert.Equal(values.Length, getRowsInfo.Length);

        File.Delete(path);
    }

    [Fact]
    public void Issue507_3_MismatchedQuoteCsv()
    {
        //Problem with multi-line when using Query func
        //https://github.com/mini-software/MiniExcel/issues/507

        var config = new CsvConfiguration
        {
            //AlwaysQuote = true,
            ReadLineBreaksWithinQuotes = true,
        };

        // create
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("A,B,C\n\"r1a: no end quote,r1b,r1c"));

        // read
        var getRowsInfo = _csvImporter.Query(stream, configuration: config).ToArray();
        Assert.Equal(2, getRowsInfo.Length);
    }

}
