<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
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
	{
		object rows = new List<Demo>(){ };
		var json = JsonConvert.SerializeObject(rows);
		Console.WriteLine(json); //[]	
	}
	{
		object rows = new List<Demo>(){ new Demo{}};
		var json = JsonConvert.SerializeObject(rows);
		Console.WriteLine(json); //[{"MyProperty1":null,"MyProperty2":null}]
	}
	{
		object rows = new List<Demo>() { new Demo { Demo2s = new List<Demo2>() {}} };
		var json = JsonConvert.SerializeObject(rows);
		Console.WriteLine(json); //[{"MyProperty1":null,"MyProperty2":null,"Demo2s":[]}]
	}
	{
		object rows = new List<Demo>() { new Demo { Demo2s = new List<Demo2>() {null } } };
		var json = JsonConvert.SerializeObject(rows);
		Console.WriteLine(json); //[{"MyProperty1":null,"MyProperty2":null,"Demo2s":[null]}]
	}
	{
		object rows = new List<Demo>() { new Demo { Demo2s = new List<Demo2>() { null, new Demo2() {} } } };
		var json = JsonConvert.SerializeObject(rows);
		Console.WriteLine(json); //[{"MyProperty1":null,"MyProperty2":null,"Demo2s":[null,{"MyProperty1":null,"MyProperty2":null}]}]
	}
	{
		object rows = new List<Demo>() { new Demo { Demo2s = new List<Demo2>() {new Demo2()} } };
		var json = JsonConvert.SerializeObject(rows);
		Console.WriteLine(json); //[{"MyProperty1":null,"MyProperty2":null,"Demo2s":[{"MyProperty1":null,"MyProperty2":null}]}]
	}
}

// You can define other methods, fields, classes and namespaces here
public class Demo
{
	public string MyProperty1 { get; set; }
	public string MyProperty2 { get; set; }
	public List<Demo2> Demo2s { get; set; }
}

public class Demo2
{
	public string MyProperty1 { get; set; }
	public string MyProperty2 { get; set; }
}