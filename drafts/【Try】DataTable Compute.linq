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
	var str = $"$=sum(D1:D{{{{$current_row}}}})";
	
	DataTable dt = new DataTable();
	dt.Columns.Add("Total", typeof(int));
	dt.Columns.Add("EmpID", typeof(int));
	dt.Rows.Add(1,4);
	dt.Rows.Add(2,5);
	dt.Rows.Add(7,5);

	// Declare an object variable.
	object sumObject;
	sumObject = dt.Compute("Sum(Total)", "EmpID = 5");
	Console.WriteLine(sumObject);
	
	Console.WriteLine(dt);
	dt.Clear();

	Console.WriteLine(dt);
	
	dt.Rows.Add(1, 4);
	dt.Rows.Add(2, 5);
	dt.Rows.Add(7, 5);
	
	Console.WriteLine(dt);
}

// You can define other methods, fields, classes and namespaces here
