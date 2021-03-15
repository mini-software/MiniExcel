<Query Kind="Program">
  <NuGetReference>AngleSharp</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>DocumentFormat.OpenXml</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <RuntimeVersion>5.0</RuntimeVersion>
</Query>

// https://github.com/shps951023/MiniExcel/issues/1
void Main()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
	var table = new DataTable();
	{
		table.Columns.Add("A", typeof(string));
		for (int i = 0; i < 10000000; i++)
			table.Rows.Add(string.Join("",Guid.NewGuid().ToString().Take(6)));
	}
	MiniExcelLibs.MiniExcel.SaveAs(path, table);
}

void Main2()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
	var table = new DataTable();
	{
		table.Columns.Add("A", typeof(string));
		for (int i = 0; i < 10000000; i++)
			table.Rows.Add("ABCDEF");
	}
	MiniExcelLibs.MiniExcel.SaveAs(path, table);
}

// You can define other methods, fields, classes and namespaces here
