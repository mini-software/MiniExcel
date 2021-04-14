<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.IO.Compression</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <Namespace>MiniExcelLibs.Attributes</Namespace>
</Query>

void Main()
{
	// ==== Excel Query ====
	{
		var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
		Console.WriteLine("==== SaveAs by strongly type ====");
		var input = JsonConvert.DeserializeObject<IEnumerable<UserAccount>>("[{\"ID\":\"78de23d2-dcb6-bd3d-ec67-c112bbc322a2\",\"Name\":\"Wade\",\"BoD\":\"2020-09-27T00:00:00\",\"Age\":5019,\"VIP\":false,\"Points\":5019.12,\"IgnoredProperty\":null},{\"ID\":\"20d3bfce-27c3-ad3e-4f70-35c81c7e8e45\",\"Name\":\"Felix\",\"BoD\":\"2020-10-25T00:00:00\",\"Age\":7028,\"VIP\":true,\"Points\":7028.46,\"IgnoredProperty\":null},{\"ID\":\"52013bf0-9aeb-48e6-e5f5-e9500afb034f\",\"Name\":\"Phelan\",\"BoD\":\"2021-10-04T00:00:00\",\"Age\":3836,\"VIP\":true,\"Points\":3835.7,\"IgnoredProperty\":null},{\"ID\":\"3b97b87c-7afe-664f-1af5-6914d313ae25\",\"Name\":\"Samuel\",\"BoD\":\"2020-06-21T00:00:00\",\"Age\":9352,\"VIP\":false,\"Points\":9351.71,\"IgnoredProperty\":null},{\"ID\":\"9a989c43-d55f-5306-0d2f-0fbafae135bb\",\"Name\":\"Raymond\",\"BoD\":\"2021-07-12T00:00:00\",\"Age\":8210,\"VIP\":true,\"Points\":8209.76,\"IgnoredProperty\":null}]");
		MiniExcel.SaveAs(path, input);
		Console.WriteLine($"File : {path}");

		{
			Console.WriteLine("==== Query strongly type ====");
			var rows = MiniExcel.Query<UserAccount>(path);
			Console.WriteLine(rows);
		}

		{
			Console.WriteLine("==== Query dynamic wihout header ====");
			var rows = MiniExcel.Query(path);
			Console.WriteLine(rows);
		}
		{
			Console.WriteLine("==== Query dynamic with useHeaderRow ====");
			var rows = MiniExcel.Query(path,useHeaderRow:true);
			Console.WriteLine(rows);
		}
		{
			Console.WriteLine("==== Query by sheetName ====");
			var rows = MiniExcel.Query(path, useHeaderRow: true,sheetName:"Sheet1");
			Console.WriteLine(rows);
		}
		{
			Console.WriteLine("==== Get All Sheets ====");
			var sheets = MiniExcel.GetSheetNames(path);
			Console.WriteLine(sheets);
		}
	}
	
	// ==== Create/Save Excel ====
	{
		var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
		Console.WriteLine("==== SaveAs by Anonymous ====");
		var input = new[] {
			new { Column1 = "MiniExcel", Column2 = 1 },
			new { Column1 = "Github", Column2 = 2}
		};
		MiniExcel.SaveAs(path, input);
		Console.WriteLine($"File : {path}");
	}
	{
		var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
		Console.WriteLine("==== SaveAs by Datatable ====");
		var table = new DataTable();
		{
			table.Columns.Add("Column1", typeof(string));
			table.Columns.Add("Column2", typeof(decimal));
			table.Rows.Add("MiniExcel", 1);
			table.Rows.Add("Github", 2);
		};
		MiniExcel.SaveAs(path, table);
		Console.WriteLine($"File : {path}");
	}
}

// You can define other methods, fields, classes and namespaces here
public class UserAccount
{
	public Guid ID { get; set; }
	public string Name { get; set; }
	public DateTime BoD { get; set; }
	[ExcelColumnName("Age")]
	public int Age2 { get; set; }
	public bool VIP { get; set; }
	public decimal Points { get; set; }
	[ExcelIgnore]
	public string IgnoredProperty { get; set; }
}
