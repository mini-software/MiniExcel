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
	var path = @"D:\git\MiniExcel\samples\csv\Test5x2.csv";
	using (var stream = File.OpenRead(path))
	using (var reader = new StreamReader(stream))
	{
		{
			var content = reader.ReadToEnd();
			Console.WriteLine("First Read:");
			Console.WriteLine(content); //result: A1...
		}
		{
			
			var content = reader.ReadToEnd();
			Console.WriteLine("Seond Read:");
			Console.WriteLine(content); //result: empty
		}
		{
			stream.Position=0;
			var content = reader.ReadToEnd();
			Console.WriteLine("After set position=0 Read:");
			Console.WriteLine(content); //result: A1...
		}
	}
}

// You can define other methods, fields, classes and namespaces here
