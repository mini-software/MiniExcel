<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <RemoveNamespace>System.Diagnostics</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Text</RemoveNamespace>
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

void Main()
{
	var path = @"D:\git\MiniExcel\samples\xlsx\TestTypeMapping.xlsx";
	Console.WriteLine(QueryAsDataTable(path));
}

// You can define other methods, fields, classes and namespaces here
public static DataTable QueryAsDataTable(string path)
{
	var rows = MiniExcel.Query(path, true);
	var dt = new DataTable();
	var first = true;
	foreach (IDictionary<string, object> row in rows)
	{
		if (first)
		{
			foreach (var key in row.Keys)
			{
				var type = row[key]?.GetType() ?? typeof(string);
				dt.Columns.Add(key, type);
			}

			first = false;
		}
		dt.Rows.Add(row.Values.ToArray());
	}
	return dt;
}