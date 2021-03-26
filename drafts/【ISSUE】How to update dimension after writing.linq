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

// https://stackoverflow.com/questions/66797421/how-replace-top-format-mark-after-streamwriter-writing
void Main()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.txt");
	using (var stream = File.Create(path))
	using (var writer = new StreamWriter(stream))
	{
		writer.WriteLine(2);
		writer.WriteLine(3);
		writer.WriteLine(4);
		
		// I want to go back to top to wirte
		stream.Position=0;
		writer.WriteLine(1);
	}
	Console.WriteLine(path);
}

// You can define other methods, fields, classes and namespaces here
