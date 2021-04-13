<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
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
	var input = new
	{
		title = "dsfsdf",
		objs = Enumerable.Range(1, 100000000).Select(s => Guid.NewGuid())
	};
	var dic = new Dictionary<string,object>();
	var type = input.GetType();
	var props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
	foreach (var prop in props)
	{
		dic.Add(prop.Name,prop.GetValue(input));
	}
	Console.WriteLine(dic);
	Console.WriteLine(dic["objs"]);	
}

// You can define other methods, fields, classes and namespaces here
