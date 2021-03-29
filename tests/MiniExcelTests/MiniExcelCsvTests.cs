using CsvHelper;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Xunit;

namespace MiniExcelLibs.Tests
{
    public class MiniExcelCsvTests
    {
	   public class Test
        {
            public string c1 { get; set; }
		  public string c2 { get; set; }
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
                var rows = stream.Query(useHeaderRow: true).ToList();
                Assert.Equal("A1", rows[0].c1);
                Assert.Equal("B1", rows[0].c2);
                Assert.Equal("A2", rows[1].c1);
                Assert.Equal("B2", rows[1].c2);
            }

		  {
			 var rows = MiniExcel.Query(path,useHeaderRow: true).ToList();
			 Assert.Equal("A1", rows[0].c1);
			 Assert.Equal("B1", rows[0].c2);
			 Assert.Equal("A2", rows[1].c1);
			 Assert.Equal("B2", rows[1].c2);
		  }
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
                var rows = stream.Query<Test>().ToList();
                Assert.Equal("A1", rows[0].c1);
                Assert.Equal("B1", rows[0].c2);
                Assert.Equal("A2", rows[1].c1);
                Assert.Equal("B2", rows[1].c2);
            }

		  {
			 var rows = MiniExcel.Query<Test>(path).ToList();
			 Assert.Equal("A1", rows[0].c1);
			 Assert.Equal("B1", rows[0].c2);
			 Assert.Equal("A2", rows[1].c1);
			 Assert.Equal("B2", rows[1].c2);
		  }
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