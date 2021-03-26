<Query Kind="Program">
  <Connection>
    <ID>2a2e3c0b-0e23-4992-bf19-66db2739e377</ID>
    <Persist>true</Persist>
    <Server>(localdb)\mssqllocaldb</Server>
    <Database>Northwind</Database>
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
	IsAssignableFromIDictionary<Dictionary<string, object>>().Dump(); // true
	IsAssignableFromIDictionary<Dictionary<int, object>>().Dump(); // true

	IsAssignableFromDapperRows<Dictionary<string, object>>().Dump(); // true
	IsAssignableFromDapperRows<Dictionary<int, object>>().Dump(); // true
	
	IsAssignableFromDapperRows(Connection.Query("select 1 id")).Dump(); // false
}

// You can define other methods, fields, classes and namespaces here
internal static bool IsAssignableFromIDictionary<T>()
{
	return typeof(IDictionary).IsAssignableFrom(typeof(T));
}

internal static bool IsAssignableFromDapperRows<T>()
{
	return typeof(IDictionary<string,object>).IsAssignableFrom(typeof(T));
}

internal static bool IsAssignableFromDapperRows<T>(IEnumerable<T> value)
{
	var type = typeof(T);
	var type2 = value.GetType();
	return typeof(IDictionary<string, object>).IsAssignableFrom(typeof(T));
}