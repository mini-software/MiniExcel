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
                stream.SaveAs(records,excelType:ExcelType.Csv);
            }

		  var content = File.ReadAllText(path);
		  File.Delete(path);
		  return content;
	   }
    }
}