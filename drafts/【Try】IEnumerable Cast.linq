<Query Kind="Program">
  <Connection>
    <ID>2a2e3c0b-0e23-4992-bf19-66db2739e377</ID>
    <Persist>true</Persist>
    <Server>(localdb)\mssqllocaldb</Server>
    <Database>tempdb</Database>
  </Connection>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>Dapper</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
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
	object d = new List<int>() {1,2};
	Console.WriteLine(d is IEnumerable); //True
	Console.WriteLine(d is Enumerable); //False
	Console.WriteLine(d is IEnumerable<object>); //false
	Console.WriteLine(d is IEnumerable<int>); //True
	Console.WriteLine(d as IEnumerable); //OK
	Console.WriteLine(d as IEnumerable<object>); //null
	
	{
		var rows0 = Connection.Query("select 1 id union all select 2"); //List<int>
		Console.WriteLine(Check(rows0)); //object
		var rows1 = Connection.Query("select 1 id union all select 2").ToList(); //List<Ojbect>
		Console.WriteLine(Check(rows1)); //object
	}
}

// You can define other methods, fields, classes and namespaces here
public Type Check<T>(IEnumerable<T> e)
{
	return typeof(T
	
	);
}